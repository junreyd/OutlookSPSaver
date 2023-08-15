/* eslint-disable react/no-unescaped-entities */
/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prettier/prettier */
/* eslint-disable react/jsx-key */
/* global console */
import * as React from "react";
import {
  ComboBox,
  IComboBox,
  IComboBoxOption,
  Spinner,
  SpinnerSize,
  Stack,
  Label,
  Text,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import { UserAgentApplication } from "msal";

export interface SitesListItem {
  title: string;
  url: string;
  siteId: string;
}

export interface DocumentListItem {
  parentId: string;
  id: string;
  name: string;
  webUrl: string;
}

export interface SitesListProps {
  items: SitesListItem[];
  companionURL: string;
  documentItems: DocumentListItem[]; // new prop for the document items
  onComboBoxChange: (
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    index?: number,
    value?: string,
    id?: string
  ) => void;
  onDocumentComboBoxChange: (
    // New callback
    event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    index?: number,
    value?: string,
    id?: string,
    parentId?: string
  ) => void;
  showErrors: boolean;
}

export interface SitesListState {
  isDocumentListLoading: boolean; // new state for the loading status of the document list
  isSiteSelected: boolean;
  isFolderSelected: boolean;
  showErrors: boolean;
}

interface DocumentComboBoxOption extends IComboBoxOption {
  parentId: string;
}

export default class Sites extends React.Component<SitesListProps, SitesListState> {
  constructor(props: SitesListProps) {
    super(props);
    this.state = {
      isDocumentListLoading: false,
      isSiteSelected: false,
      isFolderSelected: false,
      showErrors: false,
    };
    this.onComboBoxChange = this.onComboBoxChange.bind(this);
    this.onDocumentComboBoxChange = this.onDocumentComboBoxChange.bind(this); // Bind new callback
  }

  async onComboBoxChange(_event: React.FormEvent<IComboBox>, item?: IComboBoxOption): Promise<void> {
    if (item) {
      this.setState({ isSiteSelected: true, isDocumentListLoading: true });
      this.props.onComboBoxChange(_event, item, undefined, item.text, item.id);
    } else {
      this.setState({ isSiteSelected: false });
    }
  }

  // async onDocumentComboBoxChange(_event: React.FormEvent<IComboBox>, item?: IComboBoxOption): Promise<void> {
  //   if (item) {
  //     this.setState({ isFolderSelected: true });
  //     this.props.onDocumentComboBoxChange(_event, item, undefined, item.text, item.id, item.parentId);
  //   } else {
  //     this.setState({ isFolderSelected: false });
  //   }
  // }
  async onDocumentComboBoxChange(_event: React.FormEvent<IComboBox>, item?: DocumentComboBoxOption): Promise<void> {
    if (item) {
      this.setState({ isFolderSelected: true });
      this.props.onDocumentComboBoxChange(_event, item, undefined, item.text, item.id, item.parentId); // TypeScript will allow this now
    } else {
      this.setState({ isFolderSelected: false });
    }
  }

  componentDidUpdate(prevProps: SitesListProps) {
    if (prevProps.documentItems !== this.props.documentItems) {
      // when the documentItems prop changes, set the loading status to false
      this.setState({ isDocumentListLoading: false });
    }
    if (prevProps.showErrors !== this.props.showErrors) {
      this.setState({ showErrors: this.props.showErrors });
    }
  }

  render() {
    const comboBoxOptions: IComboBoxOption[] = this.props.items.map((item, index) => ({
      key: index,
      text: item.title,
      id: item.siteId,
    }));

    // map the documentItems prop to IComboBoxOption[]
    // const documentOptions: IComboBoxOption[] = this.props.documentItems.map((item, index) => ({
    //   key: index,
    //   text: item.name,
    //   id: item.id,
    //   parentId: item.parentId
    // }));

    const documentOptions: DocumentComboBoxOption[] = this.props.documentItems.map((item, index) => ({
      key: index,
      text: item.name,
      id: item.id,
      parentId: item.parentId, // No error here, since the new interface includes this property
    }));

    return (
      <main className="ms-welcome__main">
        <Text variant="mediumPlus" className="directionsText">
          Select a SharePoint site and folder to save your emails and attachments.
        </Text>
        <Stack tokens={{ childrenGap: 20 }} className="siteBody">
          <span className="companionURL">
            <a href={this.props.companionURL} target="_blank" rel="noreferrer">
              ✏️ Update list
            </a>
          </span>
          <span className="siteFieldName">Select a SharePoint site:</span>
          {this.state.showErrors && !this.state.isSiteSelected && (
            <MessageBar messageBarType={MessageBarType.error}>Please select a SharePoint site</MessageBar>
          )}
          <ComboBox
            placeholder="Choose a SharePoint site"
            options={comboBoxOptions}
            onChange={this.onComboBoxChange}
            allowFreeform={true}
            autoComplete="on"
          />
          <span className="siteFieldName">Select a folder:</span>
          {this.state.showErrors && !this.state.isFolderSelected && (
            <MessageBar messageBarType={MessageBarType.error}>Please select a folder.</MessageBar>
          )}
          {/* render a Spinner while the document list is loading */}
          {this.state.isDocumentListLoading ? (
            <Spinner size={SpinnerSize.medium} label="Loading documents..." />
          ) : (
            <>
              <ComboBox
                placeholder="Choose a folder"
                options={documentOptions}
                allowFreeform={true}
                autoComplete="on"
                onChange={this.onDocumentComboBoxChange}
              />
            </>
          )}
        </Stack>
      </main>
    );
  }
}
