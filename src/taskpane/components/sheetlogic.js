import * as React from "react";

export async function collatesheets(osheetname, ocellcolumn, ocellrow) {
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
          let sheet1rowcount = sheet1trange.rowCount;
          /* declare hash table */
          for(let j =0; j < sheet1titles.length; j++){
            let name1 = sheet1titles[i].lower().strip() /* Double check this */
            /* put the values in a hash table with the titles as the key and the
            column index as the value */
        }
        for (let x = 0; x < sheet1erange.values.length; x++){
            let email1 = sheet1erange.values[x][0].lower().strip()
            /* Put the values in a hash table with the titles as key and the row index as the value */
        }
          for (let i = 0; i < context.workbook.worksheets.length; i++){
            const sheet2 = context.workbook.worksheets[i]
            const sheet2trange = sheet2.getRange("1:10000").getUsedRange(false);
            sheet2trange.load("values");
            await context.sync();
            let sheet2titles = sheet2trange.values[0];
            let emailColumn = 0;
            let newtitles = [];
            
            for (let z = 0; z < sheet2titles.length; z++) {
                let name2 = sheet2titles[z].lower().strip()
                sheet1trange.load("columnCount")
                await context.sync();
                if (/* name is in hash table */){
                    if(name2 == "email"){
                        emailColumn = z;
                    }
                    continue;
                } else /* name not in hash table */ {
                    /* put name2 in the hash table with an index of sheet1trange.columnCount */
                    let new_cell = sheet1.getCell(0, sheet1trange.columnCount+1);
                    new_cell.load("values");
                    await context.sync();
                    new_cell.values = [[name2]];
                    newtitles.append(z); /* Check this syntax */
                    continue;
                }
            }
                
            
            temprange = sheet1.getCell(0, emailColumn)
            fullrange = temprange.getEntireColumn();
            const sheet2erange = fullrange.getUsedRange(false);
            sheet2erange.load("values");
            await context.sync(); 
            
            for(let i =0; i< sheet2erange.values.length; i++) {
                if(/* new email is in hash table */){
                    let newemailinfo = sheet2.row(rowindex);
                    let oldemailinfo = sheet1.row(rowindex);
                    oldemailinfo.load("values");
                    newemailinfo.load("values");
                    await context.sync();
                    for(let x = 0; x < newtitles.length; x++){
                        oldemailinfo.values[indexfromhash] = newemailinfo.values[newtitles[x]] /* Double check the syntax for this */
                    }
                } else /* if new email is not in the hash table*/ {
                    sheet1rowcount += 1;
                    /* Add the email to the hash table */
                    let newInfo = sheet1.row(sheet1rowcount);
                    let oldInfo = sheet2.row(currentrow);
                    newInfo.load("values");
                    oldInfo.load("values");
                    await context.sync();
                    for (let x = 0; x < sheet2trange.values.length; x++){
                        /* Search for each title's value in the hash,
                        then the hash will return the correct column in sheet1
                        then add all the values to the correct column */
                        newInfo.values[0][hashreturn] = oldInfo.values[0][newtitles[x]]
                    }

                }

            }
            /* Repeat this process for all the sheets -> Should be finished! */1
        }
      });
    } catch (error) {
        if (error.debugInfo.code = "ItemNotFound") {
          this.setState({EE2A: "Please Select a Column with Data!"});
        }
        console.error(error);
      }

}