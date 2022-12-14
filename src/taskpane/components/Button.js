import * as React from "react";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faArrowRight, faArrowDown} from '@fortawesome/free-solid-svg-icons'

const emailButton = {
    width: '100%',
    backgroundColor: '#d83b01',
    color: '#FFFFFF',
    paddingTop: '10px',
    paddingBottom: '10px',
    borderRadius: '4px',
    fontWeight: '600',
    fontSize: '14px',
    opacity: '0.85',
    cursor: 'pointer',
  }
  const sheetButton = {
    width: '100%',
    backgroundColor: '#005a9e',
    color: '#FFFFFF',
    paddingTop: '10px',
    paddingBottom: '10px',
    borderRadius: '4px',
    fontWeight: '600',
    fontSize: '14px',
    opacity: '0.85',
    cursor: 'pointer',
  }
  const iconbutton = {
    margin: '5px',
    display: 'inline',
    color: '#FFFFFF',
  }
export default class Button extends React.Component {
    render () {
        const { text, rightarrow, onClick, blue} = this.props;
        return (
        <div onClick={onClick} style={blue ? sheetButton : emailButton} classname="emailButton">
        <div style={iconbutton}> <FontAwesomeIcon icon={rightarrow ? faArrowRight : faArrowDown} /> </div>
        <span> { text } </span>
        </div>
        );
    }

}