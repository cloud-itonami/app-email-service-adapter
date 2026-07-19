import { test, expect } from '@playwright/test';

const BASE = 'https://outlook.etzhayyim.com';

test.describe('outlook.etzhayyim.com — Worker Health', () => {
  test('worker health returns ok', async ({ request }) => {
    const res = await request.get(`${BASE}/_worker/health`);
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.status).toBe('ok');
    expect(body.mode).toBe('worker-wasm');
    expect(body.app).toBe('outlook');
  });

  test('app meta endpoint is available', async ({ request }) => {
    const res = await request.get(`${BASE}/_app/meta`);
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body).toBeTruthy();
    expect(body.appId).toBe('outlook');
  });

  test('app health returns ok', async ({ request }) => {
    const res = await request.get(`${BASE}/health`);
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.status).toBe('ok');
    expect(body.app).toBe('outlook');
  });
});

test.describe('outlook.etzhayyim.com — API Endpoints', () => {
  test('GetOAuthConfig returns clientId and tenantId', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/GetOAuthConfig`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-test' },
      data: {},
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.clientId).toBeTruthy();
    expect(body.tenantId).toBeTruthy();
    expect(body.scope).toContain('Mail.Read');
  });

  test('GetConnection returns not connected for unknown user', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/GetConnection`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-unknown' },
      data: {},
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.connection.connected).toBe(false);
  });

  test('card.home returns list card', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/card.home`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-test', 'X-etzhayyim-ORG-ID': 'anon' },
      data: {},
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.contentType).toBe('application/vnd.etzhayyim.card.list');
    expect(body.payload.items.length).toBeGreaterThan(0);
  });

  test('card.compose returns form card', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/card.compose`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-test' },
      data: {},
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.contentType).toBe('application/vnd.etzhayyim.card.form');
    expect(body.payload.fields.length).toBe(3);
    expect(body.payload.action).toBe('outlook.send');
  });

  test('card.action with unknown action returns home card', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/card.action`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-test', 'X-etzhayyim-ORG-ID': 'anon' },
      data: { action: 'unknown' },
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.contentType).toBe('application/vnd.etzhayyim.card.list');
  });

  test('card.action disconnect shows confirmation card', async ({ request }) => {
    const res = await request.post(`${BASE}/xrpc/etzhayyim.outlook.v1.OutlookService/card.action`, {
      headers: { 'Content-Type': 'application/json', 'X-etzhayyim-USER-ID': 'e2e-test', 'X-etzhayyim-ORG-ID': 'anon' },
      data: { action: 'outlook.disconnect' },
    });
    expect(res.ok()).toBeTruthy();
    const body = await res.json();
    expect(body.contentType).toBe('application/vnd.etzhayyim.card.confirmation');
    expect(body.payload.destructive).toBe(true);
  });
});

test.describe('outlook.etzhayyim.com — Browser Render', () => {
  test('page loads with Outlook heading', async ({ page }) => {
    await page.goto(BASE, { waitUntil: 'networkidle' });
    await expect(page.locator('h1')).toContainText('Outlook', { timeout: 15_000 });
  });

  test('SuperAppLayout renders bottom tab bar', async ({ page }) => {
    await page.goto(BASE, { waitUntil: 'networkidle' });
    const nav = page.locator('nav.fixed, nav[class*="fixed"]');
    await expect(nav).toBeVisible({ timeout: 15_000 });
  });

  test('page shows Outlook heading and interactive state', async ({ page }) => {
    await page.goto(BASE, { waitUntil: 'networkidle' });
    // Wait for Svelte hydration + Clerk init
    await page.waitForTimeout(3000);
    // Page should have at least one button (Sign In, Connect, or Disconnect)
    const buttons = page.locator('button');
    const count = await buttons.count();
    expect(count).toBeGreaterThan(0);
  });

  test('static assets load correctly', async ({ page }) => {
    const res = await page.goto(BASE, { waitUntil: 'networkidle' });
    expect(res?.status()).toBe(200);
    // Check that JS bundles loaded (no 404s)
    const failedRequests: string[] = [];
    page.on('requestfailed', (req) => failedRequests.push(req.url()));
    await page.waitForTimeout(2000);
    const jsFailures = failedRequests.filter(u => u.includes('/_app/'));
    expect(jsFailures).toHaveLength(0);
  });
});
