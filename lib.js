// contents of lib.js
const XLSX = require("xlsx");
const Today = new Date();
// const list= ['A','B','C','D','E','F','G','H','I','J','K','L','M','N'];


function writeData(data) {
    const workbook =  XLSX.utils.book_new();
    const worksheet =  XLSX.utils.aoa_to_sheet([[]]);
    const now =  Today.getFullYear().toString() + "_" + (Today.getMonth()+1).toString() + "_" + Today.getDate().toString() ;
    let j = 0;
    XLSX.utils.sheet_add_aoa(worksheet,[["股票代號"]],{origin: 'A1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["股票名稱"]],{origin: 'B1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["價格區間"]],{origin: 'C1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["盈餘河流圖"]],{origin: 'D1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["淨值河流圖"]],{origin: 'E1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["近5年平均殖利率"]],{origin: 'F1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["近5年股利發放率"]],{origin: 'G1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["目標價格"]],{origin: 'H1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["本益比法高估價"]],{origin: 'I1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["淨值比法高估價"]],{origin: 'J1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["適用評價法"]],{origin: 'K1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["年增率"]],{origin: 'L1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["預期漲幅趴數"]],{origin: 'M1'});
    
    for (let i = 0; i < data.length; i++) {
        j ++;
        const stock = data[i].stock.split(" ");
        const river1 = data[i].river.split("/")[0];
        const river2 = data[i].river.split("/")[1];
        const dividend = data[i].dividend.split("%")[0];
        const stock_price = data[i].stock_price; 
        // const goal_price = stock_price*((80 - dividend * 3)/100+1); //目標價格:包含殖利率要超過100%
        const goal_price = stock_price*(75/100+1); //目標價格:不包含殖利率要超過70%
        const increase = Math.round((data[i].stock_applicable == '盈餘評價法' ? (data[i].stock_pe / stock_price ) - 1 : (data[i].stock_pb / stock_price) - 1 )  * 100)  ;
        
        //排除~ 
        
        if(data[i].stock_applicable == '盈餘評價法'){
          if((data[i].price != '偏低' && data[i].price != '低估') || river1 == '向下'){
            j --;
            continue;
          }
        }else 
        if(data[i].stock_applicable == '淨值評價法'){
          // if((data[i].price != '偏低' && data[i].price != '低估' && data[i].price != '合理') || river2 == '向下'){
            j --;
            continue;
          // }
        }

        //排除河流圖沒在達到目標價格
        // if(data[i].stock_applicable == '盈餘評價法'){
        //   if(goal_price > data[i].stock_pe){
        //     j --;
        //     continue;
        //   }
        // }else 
        // if(data[i].stock_applicable == '淨值評價法'){
        //   if(goal_price > data[i].stock_pb){
        //     j --;
        //     continue;
        //   }
        // }
        
        //非股息股但河流圖上限高
        // if(dividend < 4){
        //   if(data[i].stock_applicable == '盈餘評價法'){
        //     if(goal_price * 2.1 > data[i].stock_pe){
        //       j --;
        //       continue;
        //     }
        //   }else
        //   if(data[i].stock_applicable == '淨值評價法'){
        //     if(goal_price * 2.1 > data[i].stock_pb){
        //       j --;
        //       continue;
        //     }
        //   }
        // }

        XLSX.utils.sheet_add_aoa(worksheet,[[stock[1] + ".TW"]],{origin: 'A'+(j+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[stock[0]]],{origin: 'B'+(j+1)});
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].price]],{origin: 'C'+(j+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[river1]],{origin: 'D'+(j+1)});    
        XLSX.utils.sheet_add_aoa(worksheet,[[river2]],{origin: 'E'+(j+1)});
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].dividend]],{origin: 'F'+(j+1)});    
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].Payout＿Ratio]],{origin: 'G'+(j+1)});    
        XLSX.utils.sheet_add_aoa(worksheet,[[goal_price]],{origin: 'H'+(j+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].stock_pe]],{origin: 'I'+(j+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].stock_pb]],{origin: 'J'+(j+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].stock_applicable]],{origin: 'K'+(j+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].stock_YoY]],{origin: 'L'+(j+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[increase + "%"]],{origin: 'M'+(j+1)}); 
      }

    // XLSX.utils.sheet_add_aoa(worksheet);

     XLSX.utils.book_append_sheet(workbook,worksheet,"sheet1");
     XLSX.writeFileXLSX(workbook, now +  "_TW" + "_" + Date.now() + ".xlsx");
}

function writeUSData(data,j) {
    const workbook =  XLSX.utils.book_new();
    const worksheet =  XLSX.utils.aoa_to_sheet([[]]);
    const now =  Today.getFullYear().toString() + "_" + (Today.getMonth()+1).toString() + "_" + Today.getDate().toString() + "_" + j;
    
    XLSX.utils.sheet_add_aoa(worksheet,[["股票代號"]],{origin: 'A1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["目前股息"]],{origin: 'B1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["股票五年殖利率"]],{origin: 'C1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["五年P/B"]],{origin: 'D1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["每股淨值"]],{origin: 'E1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["EPS每股盈餘"]],{origin: 'F1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["EPS成長率"]],{origin: 'G1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["目標殖利率估價法"]],{origin: 'H1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["相對 P/B估價法"]],{origin: 'I1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["PEG 成長股估價法"]],{origin: 'J1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["股票價格"]],{origin: 'K1'});
    XLSX.utils.sheet_add_aoa(worksheet,[["符合項目"]],{origin: 'L1'});
    
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 今年"]],{origin: 'L1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 1年前"]],{origin: 'M1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 2年前"]],{origin: 'N1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 3年前"]],{origin: 'O1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 4年前"]],{origin: 'P1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 5年前"]],{origin: 'Q1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 6年前"]],{origin: 'R1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 7年前"]],{origin: 'S1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 8年前"]],{origin: 'T1'});
    // XLSX.utils.sheet_add_aoa(worksheet,[["EPS 9年前"]],{origin: 'U1'});

    for (let i = 0; i < data.length; i++) {
        const stock_price = data[i].dividend / (data[i].dividend_5years_avg * 1.5 * 0.01);
        let count = 0;
        if(checkNumber(round(stock_price)) != "X"){
          if(checkNumber(data[i].price) <= checkNumber(round(stock_price))){
            count++;
          }
        }

        if(checkNumber(round(data[i].pe_5years * data[i].book_value_share)) != "X"){
          if(checkNumber(data[i].price) <= checkNumber(round(data[i].pe_5years * data[i].book_value_share))){
            count++;
          }
        }

        if(checkNumber(round(data[i].eps * data[i].eps_grow)) != "X"){
          if(checkNumber(data[i].price) <= checkNumber(round(data[i].eps * data[i].eps_grow))){
            count++;
          }
        }

        XLSX.utils.sheet_add_aoa(worksheet,[[data[i].stock]],{origin: 'A'+(i+1)});      
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].dividend)]],{origin: 'B'+(i+1)});    
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].dividend_5years_avg)]],{origin: 'C'+(i+1)});   
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].pe_5years)]],{origin: 'D'+(i+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].book_value_share)]],{origin: 'E'+(i+1)});  
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].eps)]],{origin: 'F'+(i+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].eps_grow)]],{origin: 'G'+(i+1)});
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(round(stock_price))]],{origin: 'H'+(i+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(round(data[i].pe_5years * data[i].book_value_share))]],{origin: 'I'+(i+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(round(data[i].eps * data[i].eps_grow))]],{origin: 'J'+(i+1)}); 
        XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(data[i].price)]],{origin: 'K'+(i+1)});
        XLSX.utils.sheet_add_aoa(worksheet,[[count]],{origin: 'L'+(i+1)});

        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_01)]],{origin: 'L'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_02)]],{origin: 'M'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_03)]],{origin: 'N'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_04)]],{origin: 'O'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_05)]],{origin: 'P'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_06)]],{origin: 'Q'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_07)]],{origin: 'R'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_08)]],{origin: 'S'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_09)]],{origin: 'T'+(i+2)}); 
        // XLSX.utils.sheet_add_aoa(worksheet,[[checkNumber(esp_10)]],{origin: 'U'+(i+2)}); 

    }

    // XLSX.utils.sheet_add_aoa(worksheet);

     XLSX.utils.book_append_sheet(workbook,worksheet,"sheet1");
     XLSX.writeFileXLSX(workbook, now+".xlsx");
}

function round(data) {
    return Math.round(data * 100) / 100
}

function checkNumber(data){
    return Number(data) ? data : "X";
}



// let pages = {
//     goto: async where => {
//       await page.goto(where, {waitUntil: 'load', timeout: 0});
//     },
//     $: async element => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       return await page.$(element);
//     },
//     $$: async element => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       return await page.$$(element);
//     },
//     $eval: async (element, event) => {
//       await page.waitFor(2000);
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       let data = await page.$eval(
//         element,
//         (el, event) => {
//           switch (event) {
//             case 'innerText':
//               return el.innerText;
//               break;
//           }
//         },
//         event
//       );
//       if(Number(data)){
//         await page.waitFor(2000);
//         data = await page.$eval(
//           element,
//           (el, event) => {
//             switch (event) {
//               case 'innerText':
//                 return el.innerText;
//                 break;
//             }
//           },
//           event
//         );
//       }
//       return data;
//     },
//     $$eval: async (element, num, event) => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       return await page.$$eval(
//         element,
//         (els, num, event) => {
//           switch (event) {
//             case 'top':
//               return els[num].style.top;
//               break;
//             case 'left':
//               return els[num].style.left;
//               break;
//           }
//         },
//         num,
//         event
//       );
//     },
//     type: async (element, data) => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       await page.type(element, data);
//     },
//     click: async (element, options) => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//       await page.click(element, options);
//     },
//     noWaitClick: async (element, options) => {
//       await page.click(element, options);
//     },
//     press: async keyCode => {
//       await page.waitFor(2000);
//       await page.keyboard.press(keyCode);
//     },
//     check: async data => {
//       await page.waitFor(4000);
//       const el = await page.$(data);
//       return el ? el : false;
//     },
//     wait: async () => {
//       await page.waitForNavigation({waitUntil: 'networkidle2' , timeout: 200000});
//     },
//     reload: async () => {
//       await page.reload();
//     },
//     waitForLoading: async element => {
//       await page.waitForSelector(element, {
//         visible: true,
//       });
//     },
//   };
module.exports = { writeData , writeUSData };
