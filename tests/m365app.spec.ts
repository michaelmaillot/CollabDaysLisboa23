import { test, expect, Page, Response, Locator, FrameLocator } from '@playwright/test';

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

// Doesn't work on page events...

// test.skip('TeamsJS SDK init error', async ({page}) => {
//   page.on('pageerror', (error) => {
//     return error.message.indexOf('SDK initialization timed out') > -1
//   });
// });

test('App displayed in SharePoint context', async () => {
  await page.goto('https://onepointdev365.sharepoint.com/sites/CommSite-UAT/SitePages/Hello-CollabDays-Lisbon!.aspx');

  msgContext = page.locator(ENV_MESSAGE_LOCATOR_ID);

  await expect(msgContext).toHaveText(/SharePoint/);

  await expect(page.getByRole('tab', { name: 'App', exact: true })).toBeDisabled();
  await expect(page.getByRole('tab', { name: 'Pages' })).toBeDisabled();
  await expect(page.getByRole('tab', { name: 'GeoLocation' })).toBeDisabled();
  await expect(page.getByRole('tab', { name: 'AppInstallDialog' })).toBeDisabled();
  await expect(page.getByRole('tab', { name: 'Mail' })).toBeDisabled();
});

test('App displayed in Teams context', async () => {
  test.setTimeout(40000);

  // This 👇 doesn't work because Teams must init the user context (otherwise 400 HTTP error raises)
  // await page.goto('https://teams.microsoft.com/apps/3934c2bd-b5db-49a1-922d-427b51ee6823');

  await page.goto('https://teams.microsoft.com/#_');

  
  try {
    const responsePromise = waitForSPOHostingPageResponse();
    
    // This is done just in case if the app is not pinned by the user but assumed that is already installed
    await page.locator('#apps-button').click();
    await page.locator(`[apps-drag-data-id="${APP_ID}"]`).click();
    // appLocator = page.locator(`[data-tid="app-bar-${APP_ID}"]`);

    await responsePromise;

    await expect(appContext.getByRole('button', { name: 'Open Team Chat' })).toBeEnabled();

    const mailTab = appContext.getByRole('tab', { name: 'Mail' });
    await mailTab.click();
    await expect(appContext.getByRole('button', { name: 'Compose mail' })).toBeEnabled();

  } catch (error) {
    // if (error instanceof errors.TimeoutError) {
    //   await page.screenshot({ path: 'screenshot.png', fullPage: true });
    // }
  }

  getContext('iframe[name="embedded-page-container"]');

  await expect(msgContext).toHaveText(/Teams/, { timeout: 10000 });

  const mailTab = appContext.getByRole('tab', { name: 'Mail' });
  await mailTab.click();

  await expect(appContext.getByRole('button', { name: 'Compose mail' })).toBeDisabled();
});

test('App displayed in Outlook context', async () => {
  await page.goto(`https://outlook.office365.com/host/${APP_ID}`);

  await waitForSPOHostingPageResponse();
  
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

  const geoLocationTab = appContext.getByRole('tab', { name: 'GeoLocation' });
  await geoLocationTab.click();

  // Expect the button to raise an error since location permission is not included in the manifest
  const consoleEvent = page.waitForEvent('console');
  await appContext.getByRole('button', { name: 'Get current Location' }).click();

  const errorMsg = await consoleEvent;
  const geoPermissionMissing = await errorMsg.args()[0].jsonValue();
  const permissionMissingError = 'The app must specify geolocation permission in its manifest and must be granted by the user. (the user has not been asked to grant permission).'
  await expect(geoPermissionMissing.message).toBe(permissionMissingError);
});

test.afterAll(async () => {
  await page.close();
});

/**
 * Returns a response object when SPFx hosting page is loaded
 */
function waitForSPOHostingPageResponse(): Promise<Response> {

  return page.waitForResponse(response =>
    (response.url().includes('.sharepoint.com/sites/app/ClientSideAssets/')
      || response.url().includes('localhost:4321/dist'))
    && response.status() === 200
  );
}

/**
 * Sets appContext and msgContext variables for the whole test suite
 * @param frameLocator iframe locator
 */
function getContext(frameLocator: string): void {
  if (!appContext) {
    appContext = page.frameLocator(frameLocator);
  }

  if (!msgContext) {
    msgContext = appContext.locator(ENV_MESSAGE_LOCATOR_ID);
  }
}