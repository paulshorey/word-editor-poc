import * as React from "react";
import { DefaultButton } from "@fluentui/react";

import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import AddComponent from "../commands/AddComponent";
import GetFirstParagraph from "../commands/GetFirstParagraph";
import AddContentControl from "../commands/AddContentControl";
import ToggleCCDeletable from "../commands/ToggleCCDeletable";

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

  addDataElement = async () => {
    return Word.run(async (context) => {
      const contentRange = context.document.getSelection();
      const contentControl = contentRange.insertContentControl();
      contentControl.title = "title";
      contentControl.tag = "CC_TAG";
      contentControl.appearance = "Tags";
      contentControl.color = "Red";
      contentControl.cannotDelete = false;
      contentControl.cannotEdit = true;
      contentControl.appearance = "BoundingBox";

      await context.sync();

      contentControl.insertText("Hello World", "Replace");

      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
    const tagName = "CC_TAG";

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
        <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.addDataElement}
        >
          Add Data Element
        </DefaultButton>
        <br />
        <AddComponent />
        <br />
        <GetFirstParagraph />
        <br />
        <AddContentControl tagName={tagName} />
        <br />
        <ToggleCCDeletable tagName={tagName} />
      </div>
    );
  }
}
