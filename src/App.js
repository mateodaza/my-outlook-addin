import React, { Component } from 'react';
import './App.css';

class App extends Component {
  constructor(props) {
    super(props);

    this.state = {
      subject: ''
    }

    this.getAppointment = this.getAppointment.bind(this);
  }

  onSetColor() {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }

  getAppointment() {
    window.Office.context.mailbox.item.subject.getAsync(callback, (response)=>{
      this.setState({subject: response.value});
    });

    function callback(asyncResult) {
      let subject = asyncResult.value;
      console.log("Subject", subject);
      return subject;
      //this.setState({subject: subject});
    };

  }

  render() {
    return (
      <div id="content">
        <div id="content-header">
          <div className="padding">
            <h1>This is my add-in</h1>
          </div>
        </div>
        <div id="content-main">
            <button onClick={this.getAppointment.bind(this)} type="button"> Get Appointment's Body </button>
            <p> {this.state.subject} </p>
        </div>
      </div>
    );
  }
}

export default App;