import * as React from "react";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import AddComponent from "../commands/AddComponent";
import GetFirstParagraph from "../commands/GetFirstParagraph";
import AddContentControl from "../commands/AddContentControl";
import ToggleCCDeletable from "../commands/ToggleCCDeletable";
import AddDataElement from "./controls/AddDataElement";

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
      <div className="faf-bg">
        <AddDataElement />
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
