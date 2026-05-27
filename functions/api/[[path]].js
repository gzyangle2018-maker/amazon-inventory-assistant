// Cloudflare Pages Function for Amazon Inventory Assistant API
import worker from '../../src/index.js';

export async function onRequest(context) {
  const { request, env } = context;
  return worker.fetch(request, env, context);
}
