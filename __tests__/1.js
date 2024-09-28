const puppeteer = require('puppeteer');

describe('測試開啟網頁', () => {
  let browser, page;
  let url = 'https://stockman.club/systems/'
  jest.setTimeout(500000000);
  
beforeEach (async () => {
    browser = await puppeteer.launch({headless: false});
    page = await browser.newPage();
    await page.setViewport({
      width: 2000,  
      height: 900,
    });
  })
  
afterEach (() => {
    browser.close()
  })

test('開啟網站', async () => {
    await page.goto(url);
    await page.waitFor(5000);
  });
 
})