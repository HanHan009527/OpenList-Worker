import assert from "node:assert/strict"
import { test } from "node:test"
import { requestApi } from "./util"

test("OneDrive missing item is classified as WebDAV not found", async () => {
  const original = globalThis.fetch
  globalThis.fetch = async () => new Response(JSON.stringify({ error: { code: "itemNotFound", message: "The resource could not be found." } }), { status: 404 })
  try {
    await assert.rejects(requestApi({ accessToken: "test" } as any, "https://graph.microsoft.com/test", "GET"), /file not found/)
  } finally { globalThis.fetch = original }
})

test("OneDrive permission errors are not disguised as missing items", async () => {
  const original = globalThis.fetch
  globalThis.fetch = async () => new Response(JSON.stringify({ error: { code: "accessDenied", message: "Access denied" } }), { status: 403 })
  try {
    await assert.rejects(requestApi({ accessToken: "test" } as any, "https://graph.microsoft.com/test", "GET"), /Access denied/)
  } finally { globalThis.fetch = original }
})
