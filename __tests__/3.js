const puppeteer = require('puppeteer');


describe('測試開啟網頁', () => {
  let browser, page;
  jest.setTimeout(500000000);
  
  beforeEach (async () => {
    browser = await puppeteer.launch({headless: false});
    page = await browser.newPage();
    await page.setViewport({
      width: 2000,  
      height: 900,
    });
    await page.goto('https://stockman.club/user/login');
    await page.type('#email', 'a976543257@gmail.com');
    await page.waitFor(3000);
    await page.type('#password', '987654321');
    await page.waitFor(3000);
    await page.click('body > main > section > div > form > button.btn.btn_sky.form_btn');
  })
  
  afterEach (() => {
    browser.close()
  })

  test('開啟網站', async () => {
    try{
      await page.goto('https://stockman.club/systems/result');
      await page.waitFor(5000);
      const input = await page.$('#search > div.search > input');
      await input.click({ clickCount: 3 })
      await page.waitFor(3000);
      await page.type('#search > div.search > input', '1101');
      await page.waitFor(3000);
      await page.click('#search > div.search > i');
      await page.waitFor(3000);
      const stock_name = await page.$eval('div.panel > table:nth-child(2) > tbody > tr:nth-child(2) > td'
                                          ,el => el.innerText );
      console.log(stock_name);
      await page.screenshot({
        path: '1101.png',
      });
    }catch (e) {
      await page.screenshot({
        path: '1101.png',
      });
    }
  });
})