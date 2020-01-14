import * as React from "react";
import { PrimaryButton } from "office-ui-fabric-react";

export interface AppProps {
}

export interface AppState {
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
  }

  componentDidMount() {
    
  }

  click = async () => {
    return Word.run(async context => {
        const selection = context.document.getSelection();  
        const tableOoxml = selection.parentTable.getRange().getOoxml();
        await context.sync(); 
        console.log(tableOoxml.value);     
    });
  };

  render() {

    return (
      <div className="ms-welcome">
          <br/>
          <PrimaryButton
            className="ms-welcome__action"
            onClick={this.click}
          >
            Run
          </PrimaryButton>
      </div>
    );
  }
}
