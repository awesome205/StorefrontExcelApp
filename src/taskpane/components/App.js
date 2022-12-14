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
      emailbool: false,
      somethingelse: "",
      sheetbool:false,
      EE1A: "Select a Column",
      EE2A: "Select a Column",
      emailclick: null,
      emailcomplete: false,
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
    });
  };

  // click = async () => {
  //   try {
  //     await Excel.run(async (context) => {
  //       /**
  //        * Insert your Excel code here
  //        */
  //       const range = context.workbook.getSelectedRange();

  //       // Read the range address
  //       range.load("address");

  //       // Update the fill color
  //       range.format.fill.color = "yellow";

  //       await context.sync();
  //       console.log(`The range address was ${range.address}.`);
  //     });
  //   } catch (error) {
  //     console.error(error);
  //   }
  // };

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
        // Update the fill color
        usedrange.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.t`);
        this.setState({EE1A: usedrange.address, EE1CI: activeCell.columnIndex, EE1RI: activeCell.rowIndex, EE1SN: sheet.name});
      });
    } catch (error) {
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
        usedrange.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.t`);
        this.setState({EE2A: usedrange.address, EE2CI: activeCell.columnIndex, EE2RI: activeCell.rowIndex, EE2SN: sheet.name});
      });
    } catch (error) {
      console.error(error);
    }
  }

combineemails = async () => {
  this.setState({somethingelse: "The automation is running...please wait. If the sheet is large, this may take a few minutes."})
  try {
    await Excel.run(async (context) => {
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
      s1range.load("address");
      s2range.load("address");
      s1range.load("rowIndex");
      s1range.load("rowCount");
      s1range.load("values");
      s2range.load("values");
      let sheet1rows = sheet1.getRange();
      sheet1rows.load("address");
      await context.sync();
      let s1rowcount = s1range.rowCount;
      console.log((sheet1rows.address));
      console.log((s1rowcount));

      function compare(a, b) {
        // Use toUpperCase() to ignore character casing
        const bandA = a.toUpperCase();
        const bandB = b.toUpperCase();
      
       if (bandA == bandB){
        return true;
       }else{
        return false;

       }
      }

      for(let i=0; i < s1range.values.length; i++){
        if (s2range.values[i][0] == ""){
          continue;
        }
        else if(compare(s1range.values[i][0], s2range.values[i][0])){
          console.log("if #2");
          s2range.getCell(i, 0).values = [["Duplicate"]];
          uniquecount = uniquecount + 1;
        }
        else {
          let newrow = sheet1rows.getCell(s1rowcount, 0)
          newrow.copyFrom(sheet1rows.getRow(i));
          let new_cell = sheet1.getCell(s1rowcount, this.state.EE1CI);
          let copied_cell = sheet1.getCell(s1rowcount, this.state.EE2CI);
          copied_cell.load("values");
          new_cell.load("values")
          await context.sync();
          new_cell.values = copied_cell.values;
          copied_cell.values = [["New Email"]]
          s2range.getCell(i, 0).values = [["Copied To New Row"]];
          s1rowcount = s1rowcount + 1;
          uniquecount = uniquecount + 1;
        }
      }
      await context.sync();
      this.setState({somethingelse: "Task Completed", newemails: (s1rowcount - s1range.rowCount - 1), uniqueemails: uniquecount})
    });
    }catch (error){
      console.error(error);
    }
    this.setState({emailcomplete: true})
}

emailclick = () => {
  this.setState({emailbool: !this.state.emailbool})

}

sheetclick = (sheet) => {
  sheet = !sheet;
  this.setState({sheetbool: sheet})
}

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/SPM_Old-Logo_2017.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo={require("./../../../assets/SPM_Old-Logo_2017.png")}
          title={this.props.title}
          message="Welcome to the Storefront Excel App"
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
          If not, it makes a copy of the row infromation and adds a new row.
          </p>
          <div>
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>Email Column 1</p>
          <p>Please select the first column with your cursor, then select the box below.</p>
          <input onClick={this.handleInput1} style={inputstyle} value={this.state.EE1A}></input>
          </div>
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>Email Column 2</p>
          <p>Please select the second column with your cursor, then select the box below.</p>
          <input onClick={this.handleInput2} style={inputstyle} value={this.state.EE2A}></input>
          <PrimaryButton style={{marginTop: '10px'}} onClick={this.combineemails}>
            Combine Emails
          </PrimaryButton>
          {this.state.emailcomplete ? <p className="ms-font-1">Task Completed! You added {this.state.newemails} new emails.
            There are about {this.state.uniqueemails} unique emails in this sheet.</p> : 
          <p style={{ fontWeight: 'bold', fontSize: '15px', marginBottom: '0px', paddingBottom: '0px'}}>{this.state.somethingelse} </p>}
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
        <>
        <p style={pstyle2} className="ms-font-l"> 
        This program combines all the sheets in this Workbook into one. Please select one row which has a common name
        between all the sheets. The column name must be in Row 1 of the sheet. The program compares the column values between sheets.
        If they match, then the program adds additional columns to that row. If it is a unique value, the program will add an additional
        row to the new collated sheet.
          </p>
        </>
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
