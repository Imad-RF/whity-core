/**
 * @jest-environment node
 *
 * Regression guard: the catch-all API proxy MUST forward a binary response
 * body BYTE-FOR-BYTE. It previously read the upstream body with
 * `response.text()`, which decodes as UTF-8 and corrupts any non-text payload
 * — so a generated .docx download (and any other binary) arrived mangled and
 * Word refused to open it. The proxy now forwards `response.arrayBuffer()`.
 */

const BACKEND = 'http://backend.test';

jest.mock('@/lib/backend-url', () => ({
  backendUrl: () => BACKEND,
}));

jest.mock('next/headers', () => ({
  cookies: jest.fn(async () => ({
    getAll: () => [] as { name: string; value: string }[],
  })),
}));

import { GET as catchAllGET } from '@/app/api/[...path]/route';

describe('proxy binary passthrough', () => {
  const realFetch = global.fetch;

  afterEach(() => {
    global.fetch = realFetch;
    jest.clearAllMocks();
  });

  it('forwards a binary body byte-for-byte (not UTF-8 mangled)', async () => {
    // A .docx starts with the ZIP magic "PK\x03\x04" and contains arbitrary
    // bytes; several of these are invalid UTF-8 sequences that response.text()
    // would replace with U+FFFD, inflating and corrupting the file.
    const bytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0xff, 0xfe, 0x00, 0x80, 0xc3, 0x28, 0x1f]);
    global.fetch = jest.fn(async (): Promise<Response> =>
      new Response(bytes, {
        status: 200,
        headers: { 'content-type': 'application/octet-stream' },
      })
    ) as unknown as typeof fetch;

    const request = new Request('http://localhost:3000/api/bom/documents', { method: 'GET' });
    const result = await catchAllGET(request);

    expect(result.status).toBe(200);
    const out = new Uint8Array(await result.arrayBuffer());
    expect(Array.from(out)).toEqual(Array.from(bytes));
  });
});
