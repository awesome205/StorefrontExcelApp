import * as React from "react";
import PropTypes from "prop-types";

const h2style = {
  textAlign: 'center',
  margin: '10px',
}

export default class Header extends React.Component {
  render() {
    const { title, logo, message } = this.props;

    return (
      <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
        <img width="200" height="90" src={logo} alt={title} title={title} />
        <h2 style={h2style} className="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{message}</h2>
      </section>
    );
  }
}

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};
