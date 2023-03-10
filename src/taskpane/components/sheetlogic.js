import { title } from "process";
import * as React from "react";

export async function collatesheets(osheetname, ocellcolumn, ocellrow, myfunc) {
    try {
        Excel.run(async (context) => {
          /**
           * Insert your Excel code here
           */
          const sheet1 = context.workbook.worksheets.getItem(osheetname);
          const sheet1trange = sheet1.getRange("1:10000").getUsedRange(false);
          let temprange = sheet1.getCell(ocellrow, ocellcolumn)
          let fullrange = temprange.getEntireColumn();
          const sheet1erange = fullrange.getUsedRange(false);
          sheet1erange.load("values");
          sheet1erange.load("rowCount");
          sheet1trange.load("values");
          sheet1trange.load("columnCount");
          await context.sync(); 
          let sheet1titles = sheet1trange.values[0];
          let sheet1maintitle = sheet1erange.values[0][0].toLowerCase().trim();
          let sheet1rowcount = sheet1erange.rowCount;
          let sheet1columncount = sheet1trange.columnCount;
          let titlemap = new Map();
          let emailmap = new Map();
          /* declare hash table */
          let name1 = sheet1titles[0];
          for(let i = 0; i < sheet1titles.length; i++){
            name1 = sheet1titles[i].toLowerCase().trim();
            /* put the values in a hash table with the titles as the key and the
            column index as the value */
            titlemap.set(name1, i);
          }
          console.log("first for loop + column count " + sheet1trange.columnCount);
          let email1 = sheet1erange.values[0][0];
          for(let krib = 1; krib < sheet1erange.values.length; krib++){
                email1 = sheet1erange.values[krib][0];
                emailmap.set(email1, krib);
            }
        console.log("second for loop + map " + email1);
        const count = context.workbook.worksheets.getCount()
        let sheet2 = context.workbook.worksheets.getFirst();
        await context.sync();

          for (let i = 0; i < count.value-1; i++){
            sheet2.load("name");
            await context.sync();
            if(sheet2.name == osheetname){
                sheet2 = sheet2.getNextOrNullObject();
                sheet2.load("name");
                await context.sync();
            }
            // Implement check for null value at the end of the function -> if it goes to the next sheet and it is null then end the function
            console.log("Worksheet " + i);
            const sheet2trange = sheet2.getRange("1:10000").getUsedRange(false);
            sheet2trange.load("values");
            sheet2trange.load("columnCount");
            await context.sync();
            let sheet2columncount = sheet2trange.columnCount;
            let sheet2titles = sheet2trange.values[0];
            let emailColumn = 0;
            let newtitles = [];
            let localtitlemap = new Map();
            console.log("sheet 1 row count " + sheet1rowcount);
            
            for (let z = 0; z < sheet2titles.length; z++) {
                let name2 = sheet2titles[z].toLowerCase().trim();
                sheet1trange.load("columnCount");
                await context.sync();
                localtitlemap.set(name2, z);
                // localtitlemap.set(z, name2);
                if (titlemap.has(name2)){
                    if(name2 == sheet1maintitle){ // change this to the value that's being passed in.
                        emailColumn = z; // recorded so you know which one to compare to when you need to copy the values over
                    }
                } else /* name not in hash table */ {
                    titlemap.set(name2, sheet1columncount);
                    /* put name2 in the hash table with an index of sheet1trange.columnCount */
                    let new_cell = sheet1.getCell(0, sheet1columncount);
                    new_cell.load("values");
                    await context.sync();
                    new_cell.values = [[sheet2titles[z]]];
                    newtitles.push(name2); /* Check this syntax */
                    sheet1columncount += 1;
                }
            }
            console.log("Email column is " + emailColumn);
            let temprange2 = sheet2.getCell(0, emailColumn);
            let fullrange2 = temprange2.getEntireColumn();
            const sheet2erange = fullrange2.getUsedRange(false);
            sheet2erange.load("values");
            // let temprange3 = sheet2.getCell(10, 0);
            // let temp4 = temprange3.getEntireRow();
            // temp4.load("address");
            // temp4.load("values");
            await context.sync(); 
            let testrow = sheet1.getRangeByIndexes(sheet1rowcount, 0, 1, (sheet1columncount));
            testrow.load("values");
            testrow.load("address");
            await context.sync();
            console.log("TEST ROW");
            console.log(testrow.address);
            console.log(testrow.values);

            for(let ced = 1; ced < sheet2erange.values.length; ced++){
                let email2 = sheet2erange.values[i][0].toLowerCase().trim();
                if(emailmap.has(email2)){
                    console.log("GOES INSIDE FIRST IF STATEMENT 107");
                    let newemailinfo = sheet2.getRangeByIndexes(ced, 0, 1, (sheet2columncount));
                    let num = emailmap.get(email2);
                    let oldemailinfo = sheet1.getRangeByIndexes(num, 0, 1, (sheet1columncount));
                    oldemailinfo.load("values");
                    newemailinfo.load("values");
                    await context.sync();
                    console.log(oldemailinfo.values);
                    console.log(newemailinfo.values);
                    for(let xo = 0; xo < newtitles.length; xo++){
                        oldemailinfo.values[0][titlemap.get(newtitles[xo])] = newemailinfo.values[0][localtitlemap.get(newtitles[xo])]; /* Double check the syntax for this */
                    }
                    await context.sync();
                } else {
                    console.log("other implementation 2");
                    for (let [key, value] of localtitlemap){
                        let copycell = sheet2.getCell(ced, value);
                        let destinationcell = sheet1.getCell(sheet1rowcount, titlemap.get(key));
                        destinationcell.load("values");
                        copycell.load("values");
                        await context.sync();
                        destinationcell.values = copycell.values;

                    }
                    // for(let column = 0; column < sheet2columncount; column++) {
                    //     let copycell = sheet2.getCell(ced, column);
                    //     let destinationcell = sheet1.getCell(sheet1rowcount, titlemap.get(localtitlemap.get(column)));
                    //     destinationcell.load("values");
                    //     copycell.load("values");
                    //     await context.sync();
                    //     destinationcell.values = copycell.values;
                    // }
                    sheet1rowcount += 1;
                    await context.sync();

                }
            }

            console.log("Emails were copied");


            sheet2 = sheet2.getNextOrNullObject();
            await context.sync();
            if (sheet2.isNullObject){
                console.log("Did it breaK?");
                break;
            }
            
        }
      }).then((res) => {
        myfunc();
      });
    } catch (error) {
        if (error.debugInfo.code = "ItemNotFound") {
          this.setState({EE2A: "Please Select a Column with Data!"});
        }
        console.error(error);
      }

}