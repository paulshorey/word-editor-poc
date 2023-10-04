import * as React from "react";
import { DefaultButton } from "@fluentui/react";

import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import AddComponent from "../commands/AddComponent";
import GetFirstParagraph from "../commands/GetFirstParagraph";
import AddContentControl from "../commands/AddContentControl";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
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
  }

  click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const doc = context.document;
      const originalRange = doc.getSelection();
      const paragraph = originalRange.insertText("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "red";

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
          Run
        </DefaultButton>
        <AddComponent />
        <GetFirstParagraph />
        <AddContentControl />
      </div>
    );
  }
}
