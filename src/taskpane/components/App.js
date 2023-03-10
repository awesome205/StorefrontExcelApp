/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prettier/prettier */
import * as React from "react";
import PropTypes from "prop-types";
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import Header from "./Header";
// import HeroList from "./HeroList";
import Progress from "./Progress";
import Button from "./Button"
import { PrimaryButton } from "@fluentui/react";
import { Spinner, SpinnerSize } from "@fluentui/react";
import { collatesheets } from "./sheetlogic";



/* global console, Excel, require */

const sheetbutton = {
  marginTop: '20px',
}
const pstyle = {
  fontSize: '15px',
  marginBottom: '10px',
}
const pstyle2 = {
  fontSize: '15px',
  marginBottom: '20px',
}
const divstyle = {
  marginLeft: '8px',
  paddingBottom: '10px',
}
const inputstyle = {
  width: '90%',
  padding: '10px',
  boxSizing: 'border-box',

}
const innercontent = {
  margin:'7.5px',
}



export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      emailbool: true,
      somethingelse: "",
      sheetbool:true,
      EE1A: "Select a Column",  
      EE2A: "Select a Column",
      emailclick: null,
      emailcomplete: false,
      columntitle: [],
      sheetload: "",
    };
  }


  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
      columntitle: [{sheet: "", row: 0, column: 0, value: "Please Select a Column Name"}]
    });
  };

  handleInput3 = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const range = context.workbook.getSelectedRange();
        let activeCell = context.workbook.getActiveCell();
      
        let entireRange = activeCell.getEntireColumn();
        const usedrange = entireRange.getUsedRange(false);

        // Read the range address
        usedrange.load("address");
        range.load("address");
        activeCell.load("columnIndex");
        activeCell.load("rowIndex");
        usedrange.load("values");
        // Update the fill color
        // usedrange.format.fill.color = "yellow";

        await context.sync();
        console.log(usedrange.values);
        console.log(`The range address was ${range.address}.t`);
        this.setState({EE1A: usedrange.address, EE1CI: activeCell.columnIndex, EE1RI: activeCell.rowIndex, EE1SN: sheet.name});
      });
    } catch (error) {
      if (error.debugInfo.code = "ItemNotFound") {
        this.setState({EE1A: "Please Select a Column with Data!"});
      }
      console.error(error);
      

    }
  }
  handleInput1 = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const range = context.workbook.getSelectedRange();
        let activeCell = context.workbook.getActiveCell();
      
        let entireRange = activeCell.getEntireColumn();
        const usedrange = entireRange.getUsedRange(false);

        // Read the range address
        usedrange.load("address");
        range.load("address");
        activeCell.load("columnIndex");
        activeCell.load("rowIndex");
        usedrange.load("values");
        // Update the fill color
        // usedrange.format.fill.color = "yellow";

        await context.sync();
        console.log(usedrange.values);
        console.log(`The range address was ${range.address}.t`);
        this.setState({EE1A: usedrange.address, EE1CI: activeCell.columnIndex, EE1RI: activeCell.rowIndex, EE1SN: sheet.name});
      });
    } catch (error) {
      if (error.debugInfo.code = "ItemNotFound") {
        this.setState({EE1A: "Please Select a Column with Data!"});
      }
      console.error(error);
      

    }
  }
  handleInput2 = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const range = context.workbook.getSelectedRange();
        let activeCell = context.workbook.getActiveCell();
      
        let entireRange = activeCell.getEntireColumn();
        const usedrange = entireRange.getUsedRange(false);

        // Read the range address
        usedrange.load("address");
        range.load("address");
        activeCell.load("columnIndex");
        activeCell.load("rowIndex");
        // Update the fill color
        // usedrange.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.t`);
        this.setState({EE2A: usedrange.address, EE2CI: activeCell.columnIndex, EE2RI: activeCell.rowIndex, EE2SN: sheet.name});
      });
    } catch (error) {
      if (error.debugInfo.code = "ItemNotFound") {
        this.setState({EE2A: "Please Select a Column with Data!"});
      }
      console.error(error);
    }
  }

combineemails = async () => {
  if (this.state.EE1A == this.state.EE2A){
    this.setState({somethingelse: "Please select two unique columns and try again."})
    return
  }
  if (this.state.emailcomplete){
    this.setState({emailcomplete: false});
  }
  this.setState({somethingelse: "The automation is running."})
  try {
     Excel.run(async (context) => {
      console.log("the function ran")
      //Need to get the range with the last used cell in Excel -> then operate over these
      //cells 
      let uniquecount = 0;
      let sheet1 = context.workbook.worksheets.getItem(this.state.EE1SN);
      let sheet2 = context.workbook.worksheets.getItem(this.state.EE2SN);
      let sheet1firstcell = sheet1.getCell(this.state.EE1RI, this.state.EE1CI);
      let sheet2firstcell = sheet2.getCell(this.state.EE2RI, this.state.EE2CI);
      let i1range = sheet1firstcell.getEntireColumn();
      let i2range = sheet2firstcell.getEntireColumn();
      let s1range = i1range.getUsedRange(false);
      let s2range = i2range.getUsedRange(false);
      let s1lastcell = s1range.getLastCell();
      s1lastcell.load("rowIndex");
      s1range.load("address");
      s2range.load("address");
      s1range.load("rowCount");
      s1range.load("rowIndex")
      s1range.load("values");
      s2range.load("values");
      let sheet1rows = sheet1.getRange();
      sheet1rows.load("address");
      await context.sync();
      console.log("the range is " + s1range.address);
      console.log("the 2nd range is " + s2range.address);
      let s1rowcount = s1lastcell.rowIndex;

      function compare(a, b) {
        // Use toUpperCase() to ignore character casing
        const bandA = a.toUpperCase();
        const bandB = b.toUpperCase();
      
       if (bandA == bandB){
        return true;
       } else{
        return false;

       }
      }

      for(let i=0; i < s1range.values.length; i++){
        if (s2range.values[i][0] == ""){
          if (s1range.values[i][0] != ""){
            uniquecount = uniquecount + 1;
          }
          continue;
        } else if (s1range.values[i][0] == ""){
          s1range.getCell(i, 0).values = [s2range.values[i]];
          s2range.getCell(i, 0).values = [["Copied to Left"]];
          uniquecount = uniquecount + 1;
        }
        else if(compare(s1range.values[i][0], s2range.values[i][0])){
          s2range.getCell(i, 0).values = [["Duplicate"]];
          uniquecount = uniquecount + 1;
        }
        else {
          s1rowcount = s1rowcount + 1;
          let newrow = sheet1rows.getCell(s1rowcount, 0)
          let cell_row = s1range.getCell(i, 0);
          cell_row.load("rowIndex");
          await context.sync();
          newrow.copyFrom(sheet1rows.getRow(cell_row.rowIndex));
          let new_cell = sheet1.getCell(s1rowcount, this.state.EE1CI);
          let copied_cell = sheet1.getCell(s1rowcount, this.state.EE2CI);
          copied_cell.load("values");
          new_cell.load("values")
          await context.sync();
          new_cell.values = copied_cell.values;
          copied_cell.values = [["New Email"]]
          s2range.getCell(i, 0).values = [["Copied To New Row"]];
          uniquecount = uniquecount + 2;
        }
      }
      await context.sync();
      this.setState({somethingelse: "Task Completed", newemails: ((s1rowcount - s1lastcell.rowIndex)), uniqueemails: uniquecount})
    }).then((res) => {
      this.setState({emailcomplete: true});
    });
    }catch (error){
      console.error(error);
    }
}

emailclick = () => {
  this.setState({emailbool: !this.state.emailbool})

}

sheetclick = (sheet) => {
  sheet = !sheet;
  this.setState({sheetbool: sheet})
}

handlesheetInput = async () => {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      let activeCell = context.workbook.getActiveCell();

      // Read the range address
      activeCell.load("values")
      activeCell.load("address"); 
      activeCell.load("rowIndex");
      activeCell.load("columnIndex");
    
      await context.sync();
  
      console.log(`Worksheet name: ${sheet.name}`);
  
      this.setState({columntitle: [{sheet: sheet.name, row: activeCell.rowIndex, column: activeCell.columnIndex, value: activeCell.values[0][0]}]});
    });
  } catch (error) {
    if (error.debugInfo.code = "ItemNotFound") {
      this.setState({EE2A: "Please Select a Column with Data!"});
    }
    console.error(error);
  }
}
finishsheet = () => {
  this.setState({sheetload: "The collating is complete!"});
}
csheetclick = () => {
  this.setState({sheetload: "Please wait the sheets are being collated."})
  // collatesheets(this.state.columntitle[0].sheet, this.state.columntitle[0].column, this.state.columntitle[0].row, this.finishsheet);
  let osheetname = this.state.columntitle[0].sheet;
  let ocellcolumn = this.state.columntitle[0].column;
  let ocellrow = this.state.columntitle[0].row;

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
      let localtitlemap = new Map();
      titlemap.clear();
      emailmap.clear();
      localtitlemap.clear();
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
            console.log(email1);
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
        console.log(count);
        console.log("Worksheet " + i + sheet2.name);
        const sheet2trange = sheet2.getRange("1:10000").getUsedRange(false);
        sheet2trange.load("values");
        sheet2trange.load("columnCount");
        await context.sync();
        let sheet2columncount = sheet2trange.columnCount;
        let sheet2titles = sheet2trange.values[0];
        let emailColumn = 0;
        let newtitles = [];
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
        console.log(newtitles);
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
            let email2 = sheet2erange.values[ced][0].toLowerCase().trim();
            if(emailmap.has(email2)){
                console.log("GOES INSIDE FIRST IF STATEMENT 107");
                for(let xo = 0; xo < newtitles.length; xo++){
                  let newemailinfo = sheet2.getCell(ced, localtitlemap.get(newtitles[xo]));
                  let oldemailinfo = sheet1.getCell(emailmap.get(email2), titlemap.get(newtitles[xo]));
                  oldemailinfo.load("values");
                  newemailinfo.load("values");
                  await context.sync();
                  oldemailinfo = newemailinfo;
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
                emailmap.set(email2, sheet1rowcount);
                sheet1rowcount += 1;
                await context.sync();
            }
        }

        console.log("Emails were copied");


        // sheet2 = sheet2.getNextOrNullObject();
        // await context.sync();
        // if (sheet2.isNullObject){
        //     console.log("Did it breaK?");
            break;
        // }
        
    }
  }).then((res) => {
    this.finishsheet();
  });
} catch (error) {
    if (error.debugInfo.code = "ItemNotFound") {
      this.setState({EE2A: "Please Select a Column with Data!"});
    }
    console.error(error);
  }







}



  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/EmailSheet.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          message="Welcome to the Email & Sheet Automation App"
        />
        {/* <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}> */}
        <div style={innercontent}>
        <p className="ms-font-l text-1">
        Select one of the dropdowns below to get started.
        </p>
        <Button text="Email Automation" rightarrow={this.state.emailbool} onClick={() => this.emailclick()} />
       
        { this.state.emailbool ? (
          <p style={pstyle} className="ms-font-l"> Your solution for combining and de-duping emails.
          </p> ):(
          <>
          <div style={divstyle}>
          <p style={pstyle} className="ms-font-l"> This program will take two columns as input.
          It checks whether the emails are duplicates. If so, it deletes one of the emails.
          If not, it makes a copy of the row information and adds a new row.
          </p>
          <div>
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>Email Column 1</p>
          <p>Please select the first column with your cursor <strong> first</strong>, then select the box below.</p>
          <input onClick={this.handleInput1} className={this.state.EE1A == "Please Select a Column with Data!" ? "Iteminputred" : "Iteminput"} style={inputstyle} value={this.state.EE1A}></input>
          </div>
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>Email Column 2</p>
          <p>Please select the second column with your cursor, then select the box below.</p>
          <input onClick={this.handleInput2} className={this.state.EE2A == "Please Select a Column with Data!" ? "Iteminputred" : "Iteminput"} style={inputstyle} value={this.state.EE2A}></input>
          <PrimaryButton style={{marginTop: '10px'}} onClick={this.combineemails}>
            Combine Emails
          </PrimaryButton>
          {this.state.emailcomplete ? <p className="ms-font-1">Task Completed! You added {this.state.newemails} new emails.
            There are about {this.state.uniqueemails} unique emails in this sheet.</p> : 
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}> {this.state.somethingelse == "The automation is running." ? <Spinner size={SpinnerSize.small} label={"The automation is running...please wait. If the sheet is large, this may take a few minutes."} /> : null} </p>}
          </div>
          </>
        )}
        <div style={sheetbutton}>
        <Button blue={true} text="Sheet Automation" rightarrow={this.state.sheetbool} onClick={() => this.sheetclick(this.state.sheetbool)} />
        </div>
        {this.state.sheetbool ? (
          <p style={pstyle} className="ms-font-l"> Your solution for collating all sheets into one.
          </p>
        ) : 
        <div style={divstyle}>
        <p style={pstyle2} className="ms-font-l"> 
        This program combines all the sheets in this workbook into one. Please select one row which has a common name
        between all the sheets. The column name must be in Row 1 of the sheet. The program compares the column values between sheets.
        If they match, then the program adds additional columns to that row. If it is a unique value, the program will add an additional
        row to the new collated sheet.
          </p>
          <div style={{paddingBottom: '10px'}}>
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>Common Column</p>
          <p>Please select the name of the column which is common among all sheets.</p>
          <input onClick={this.handlesheetInput} style={inputstyle} value={this.state.columntitle[0].value}></input>
          </div>
          {this.state.columntitle[0].value != "Please Select a Column Name" ? 
          (<div>
          <p>You selected the {this.state.columntitle[0].value} column. All sheets <strong>must</strong> have this
          column title in order to be compared and combined. Press below to collate the sheets.</p>
          <PrimaryButton style={{marginTop: '5px'}} onClick={this.csheetclick}>
            Collate Sheets </PrimaryButton>
          </div>
          ) : null}
          {this.state.sheetload == "The collating is complete!" ? <p className="ms-font-1">Task Completed! The collating is finished.</p> : 
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}> {this.state.sheetload == "Please wait the sheets are being collated." ? <Spinner size={SpinnerSize.small} label={"The automation is running...please wait. If there are numerous sheets, this may take a few minutes."} /> : null} </p>}
          
        </div>
        
        }
        {/* </HeroList> */}
        </div>
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};