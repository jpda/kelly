import { DefaultPalette, Icon, IStackStyles, IStackTokens, Stack } from "office-ui-fabric-react";
import * as React from "react";

export interface ConversationListItem {
  icon: string;
  primaryText: string;
}

export interface ConversationListProps {
  items: ConversationListItem[];
}

// Styles definition
const containerStackTokens: IStackTokens = { childrenGap: 5 };

const clickableStackTokens: IStackTokens = {
  padding: 10,
};

const stackStyles: IStackStyles = {
  root: {
    background: DefaultPalette.themeTertiary,
  },
};

async function _onConversationClick(sheetName: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(sheetName);
      sheet.activate();
      sheet.load("name");
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export default class ConversationList extends React.Component<ConversationListProps> {
  render() {
    const { children, items } = this.props;

    const listItems = items.map((item, index) => (
      <Stack
        onClick={() => _onConversationClick(item.primaryText)}
        styles={stackStyles}
        tokens={clickableStackTokens}
        key={index}
      >
        <Icon iconName={item.icon} />
        <span>{item.primaryText}</span>
      </Stack>
    ));
    return (
      <Stack tokens={containerStackTokens}>
        {children}
        {listItems}
      </Stack>
    );
  }
}
