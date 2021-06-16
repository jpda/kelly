import { IconButton, IStackProps, Panel, PrimaryButton, Stack, TextField } from "office-ui-fabric-react";
import * as React from "react";
import { useBoolean } from "@fluentui/react-hooks";

export interface SettingsProps {}

const sidIconProps = { iconName: "Lock" };
const keyIconProps = { iconName: "Key" };
const phoneIconProps = { iconName: "Phone" };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};

export interface IAccountState {
  sid: string;
  key: string;
  phone: string;
}

const explanation = "Set your twilio credentials here";

const Settings: React.FC<SettingsProps> = ({}) => {
  const [sid, setSid] = React.useState("");
  const [key, setKey] = React.useState("");
  const [phone, setPhone] = React.useState("");
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);

  function handleSetSid(event) {
    setSid(event.target.value);
  }
  function handleSetKey(event) {
    setKey(event.target.value);
  }
  function handleSetPhone(event) {
    setPhone(event.target.value);
  }

  return (
    <div>
      {explanation}
      <IconButton iconProps={{ iconName: "Settings" }} onClick={openPanel} />
      <Panel
        isLightDismiss
        isOpen={isOpen}
        onDismiss={dismissPanel}
        closeButtonAriaLabel="Close"
        headerText="Light dismiss panel"
      >
        <p>{explanation}</p>
        <Stack {...columnProps}>
          <TextField
            label="Account SID"
            iconProps={sidIconProps}
            placeholder="Enter your Twilio SID here"
            onChange={handleSetSid}
            value={sid}
          />
          <TextField
            label="Account key"
            type="password"
            canRevealPassword
            iconProps={keyIconProps}
            onChange={handleSetKey}
            value={key}
          />
          <TextField
            label="Phone number to send from"
            mask="+1 999 999 9999"
            iconProps={phoneIconProps}
            onChange={handleSetPhone}
            value={phone}
          />
        </Stack>
        <PrimaryButton onClick={() => _saveSettings(sid, key, phone)}>Save settings</PrimaryButton>
      </Panel>
    </div>
  );
};

function _saveSettings(sid, key, phone): void {
  var a: IAccountState = { sid: sid, key: key, phone: phone };
  localStorage.setItem("accountInfo", JSON.stringify(a));
}

export default Settings;
