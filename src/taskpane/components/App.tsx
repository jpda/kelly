import * as React from "react";
import Header from "./Header";
import ConversationList, { ConversationListItem } from "./ConversationList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
import "../../../assets/kelly.png";
import Settings from "./Settings";
/* global console, Excel  */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: ConversationListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.loadConversations();
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  };

  async loadConversations() {
    try {
      await Excel.run(async (context) => {
        var sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        var conversations: ConversationListItem[] = [];

        sheets.items.forEach(function (sheet) {
          conversations.push({
            icon: "Message",
            primaryText: sheet.name,
          });
        });

        this.setState({
          listItems: conversations,
        });
      });
    } catch (err) {
      console.error(err);
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <>
        <div className="ms-welcome">
          <Settings />
          <Header logo="assets/kelly-circle.png" title={this.props.title} message="Messages" />
        </div>
        <ConversationList items={this.state.listItems} />
      </>
    );
  }
}
