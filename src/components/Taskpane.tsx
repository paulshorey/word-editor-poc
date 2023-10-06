import * as React from "react";
import { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import AddComponent from "@src/components/commands/AddComponent";
import GetFirstParagraph from "@src/components/commands/GetFirstParagraph";
import AddContentControl from "@src/components/commands/AddContentControl";
import ToggleCCDeletable from "@src/components/commands/ToggleCCDeletable";
import AddDataElement from "./controls/AddDataElement";
import PrepareCC4Save from "@src/components/commands/PrepareCC4Save";
import Scroll2LastComponent from "@src/components/commands/Scroll2LastComponent";

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
          logo={require("../../public/assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="faf-bg" style={{ display: "flex", flexDirection: "column", gap: "10px", padding: "0 18px" }}>
        <AddDataElement />
        <AddComponent />
        <GetFirstParagraph />
        <AddContentControl tagName={tagName} />
        <ToggleCCDeletable tagName={tagName} />
        <PrepareCC4Save />
        <Scroll2LastComponent />
      </div>
    );
  }
}
