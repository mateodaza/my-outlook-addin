import React, { Component } from 'react';
import './App.css';

class App extends Component {
  constructor(props) {
    super(props);

    this.onSetColor = this.onSetColor.bind(this);
  }

  onSetColor() {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }

  render() {
    return (
      <div id="content">
        <div id="content-header">
          <div className="padding">
              <h1>This is my add-in</h1>
          </div>
        </div>
      </div>
    );
  }
}

export default App;