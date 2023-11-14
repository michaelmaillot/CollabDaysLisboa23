import { test, expect, Page, Response, Locator, FrameLocator, errors } from '@playwright/test';

const APP_ID = "3934c2bd-b5db-49a1-922d-427b51ee6823";
const ENV_MESSAGE_LOCATOR_ID = "#envMessage";
const OUTLOOK_M365_FRAME = 'iframe[data-tid="app-host-iframe"]';

let page: Page;
let msgContext: Locator;
let appContext: FrameLocator;

test.use({
  viewport: { width: 1920, height: 1080 },
});

test.beforeAll(async ({ browser }) => {
  page = await browser.newPage();
});

// test.skip('TeamsJS SDK init error', async ({page}) => {
//   page.on('pageerror', (error) => {
//     return error.message.indexOf('SDK initialization timed out') > -1
//   });
// });

test('App displayed in Teams context', async () => {
  test.setTimeout(40000);

  // await page.goto('https://teams.microsoft.com/_#/conversations/G%C3%A9n%C3%A9ral?threadId=19:76388c600cba4a64bf8ba14004b1e0b3@thread.skype&ctx=channel');
  // await page.locator('[data-tid="app-bar-14d6962d-6eeb-4f48-8890-de55454bb136"]').click();
  // await expect(page.getByRole('heading', { name: 'Feed' })).toBeVisible();

  await page.goto('https://teams.microsoft.com/#_');

  // If you have a policy that addds automatically the custom app as pinned 👇
  try {
    const appLocator = page.locator(`[data-tid="app-bar-${APP_ID}"]`);
    const responsePromise = waitForSPOHostingPageResponse();

    await appLocator.click();
    await responsePromise;
  } catch (error) {
    // if (error instanceof errors.TimeoutError) {
    //   await page.screenshot({ path: 'screenshot.png', fullPage: true });
    // }
  }

  getContext('iframe[name="embedded-page-container"]');

  await expect(msgContext).toHaveText(/Teams/);

  const mailTab = appContext.getByRole('tab', { name: 'Mail' });
  await mailTab.click();

  await expect(appContext.getByRole('button', { name: 'Compose mail' })).toBeDisabled();
});

test('App displayed in Outlook context', async () => {
  await page.goto(`https://outlook.office365.com/host/${APP_ID}`);

  try {
    await waitForSPOHostingPageResponse();
  } catch (error) {
    // if (error instanceof errors.TimeoutError) {
    //   await page.screenshot({ path: 'screenshot.png', fullPage: true });
    // }
  }

  getContext(OUTLOOK_M365_FRAME);

  await expect(msgContext).toHaveText(/Outlook/, { timeout: 10000 });

  // const heading = page.frameLocator('iframe[data-tid="app-host-iframe"]').getByRole('heading', { name: 'Well done, Megan Bowen!' });
  // await heading;
  // await expect(heading).toBeVisible();

  const appInstallDialogTab = appContext.getByRole('tab', { name: 'AppInstallDialog' });
  await appInstallDialogTab.click();
  await expect(appContext.getByRole('button', { name: 'Open App Info' })).toBeDisabled();

  const mailTab = appContext.getByRole('tab', { name: 'Mail' });
  await mailTab.click();
  await expect(appContext.getByRole('button', { name: 'Compose mail' })).toBeEnabled();
});

test('App displayed in M365 context', async () => {
  // page.on('console', async msg => {
  //   if (msg.type() === "error" && msg.text().indexOf('Error: SDK initialization timed out') > 0) {
  //     test.skip();
  //   } 
  // });
  
  await page.goto(`https://www.microsoft365.com/m365apps/${APP_ID}?auth=2`);

  // await expect(page.getByTestId('app-header-title-button')).toHaveText('HelloM365');
  await waitForSPOHostingPageResponse();  

  getContext(OUTLOOK_M365_FRAME);

  await expect(msgContext).toHaveText(/microsoft365.com/);
});

test.afterAll(async () => {
  await page.close();
});

function waitForSPOHostingPageResponse(): Promise<Response> {

  return page.waitForResponse(response =>
    (response.url().indexOf('.sharepoint.com/sites/app/ClientSideAssets/') > 0
      || response.url().indexOf('localhost:4321/dist') > 0)
    && response.status() === 200
  );
}

function getContext(frameLocator: string): void {
  if (!appContext) {
    appContext = page.frameLocator(frameLocator);
  }

  if (!msgContext) {
    msgContext = appContext.locator(ENV_MESSAGE_LOCATOR_ID);
  }
}