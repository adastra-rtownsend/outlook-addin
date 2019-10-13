import * as React from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { ContextualMenu } from 'office-ui-fabric-react/lib/ContextualMenu';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Slider } from 'office-ui-fabric-react/lib/Slider';
import { getDefaultSettings } from '../../utilities/config';

export interface ISettingsDialogStates {
  hideDialog: boolean;
  showIntro: number;
  useSampleData: boolean;
  apiBasePath: string;
}

export default class SettingsDialog extends React.Component<{}, ISettingsDialogStates> {
  public state: ISettingsDialogStates = {
    hideDialog: true,
    showIntro: Office.context.roamingSettings.get('adastra.showIntro'),
    useSampleData: Office.context.roamingSettings.get('adastra.useSampleData'),
    apiBasePath: Office.context.roamingSettings.get('adastra.apiBasePath'),
  };
  
  private _dragOptions = {
    moveMenuItemText: 'Move',
    closeMenuItemText: 'Close',
    menu: ContextualMenu
  };

  // this is a very non-react'y way to do this, but since this settings dialog is likely 
  // temporary, we'll roll with it
  public showDialog() {
    this._showDialog();
  };

  private _restoreDefaults() {
    const defaults = getDefaultSettings();
    Office.context.roamingSettings.set('adastra.showIntro', defaults.showIntro);
    Office.context.roamingSettings.set('adastra.useSampleData', defaults.useSampleData);
    Office.context.roamingSettings.set('adastra.apiBasePath', defaults.apiBasePath);
    Office.context.roamingSettings.saveAsync();
  };

  private _applySetting(setting, value) {
    Office.context.roamingSettings.set(setting, value);
    Office.context.roamingSettings.saveAsync();
  };

  private _onToggleSampleData = ({}, checked: boolean) => {
    this.setState({ useSampleData: checked });
    this._applySetting('adastra.useSampleData', checked);
  };

  private _onSetShowWelcomeScreen = (value: number) => {
    this.setState({ showIntro: value });
    this._applySetting('adastra.useSampleData', value);
  };

  private _onSetUrl = ({}, newValue?: string) => {
    const val = newValue || '';
    this.setState({ apiBasePath: val});
    this._applySetting('adastra.apiBasePath', val);
  };

  private _getShowIntroHintText(value) {
    let msg = '';
    switch (value) {
      case 3: 
        msg = 'Always';
        break;
      case 2: 
        msg = 'Next time';
        break;
      default: 
        msg = 'Never';
        break;
    }
    return msg;
  }

  public render() {
    const { hideDialog } = this.state;
    return (
      <Dialog
        hidden={hideDialog}
        onDismiss={this._closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Ad Astra Settings',
          subText: 'Here you can change settings for the Ad Astra add-in'
        }}
        modalProps={{
          isBlocking: false,
          isModeless: true,
          dragOptions: this._dragOptions
        }}
      >
        <Slider label="Show welcome screen" 
          min={1} max={3} step={1} 
          valueFormat={(value: number) => this._getShowIntroHintText(value)} 
          value={this.state.showIntro}
          onChange={(value: number) => this._onSetShowWelcomeScreen(value)}
          showValue={true} 
        />
        <Toggle 
          label="Use sample room data" 
          checked={this.state.useSampleData} 
          onChange={this._onToggleSampleData} inlineLabel 
        /> 
        <TextField label="URL for API Bridge" 
          value={this.state.apiBasePath} 
          onChange={this._onSetUrl}
        />          
        <Link onClick={this._restoreDefaults}>Reset All Settings</Link>
        <DialogFooter>
          {/* <PrimaryButton onClick={this._saveAndClose} text="Save" /> */}
          {/* <DefaultButton onClick={this._closeDialog} text="Close" /> */}
        </DialogFooter>
      </Dialog>
    );
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  };

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };
}
