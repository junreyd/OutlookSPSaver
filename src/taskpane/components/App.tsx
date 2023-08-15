/* eslint-disable no-useless-escape */
/* eslint-disable no-control-regex */
/* eslint-disable office-addins/no-office-initialize */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Sites, { SitesListItem, DocumentListItem } from "./Sites";
import Progress from "./Progress";
import { UserAgentApplication, AuthResponse } from "msal";
import { ComboBox, IComboBox, IComboBoxOption, Spinner, SpinnerSize } from "@fluentui/react";
import Footer from "./Footer";
import { AuthContext } from "./AuthContext";
// import { Client } from "@microsoft/microsoft-graph-client";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  site: SitesListItem[];
  library: DocumentListItem[];
  listItems: HeroListItem[];
  emailAddress: string;
  isCheckboxChecked: boolean;
  siteValue: string;
  libraryId: string;
  parentLibraryId: string;
  showErrors: boolean;
  selectedSite: string;
  selectedLibrary: string;
  isLoading: boolean;
  showSuccess: boolean;
  newFolderChecked: boolean;
  companionAppurl: string;
  hasAttachment: boolean;
}

export default class App extends React.Component<AppProps, AppState> {
  static contextType = AuthContext;
  constructor(props: AppProps) {
    super(props);
    this.state = {
      site: [],
      library: [],
      listItems: [],
      emailAddress: "",
      isCheckboxChecked: false,
      siteValue: "",
      libraryId: "",
      parentLibraryId: "",
      showErrors: false,
      selectedSite: null,
      selectedLibrary: null,
      isLoading: false,
      showSuccess: false,
      newFolderChecked: false,
      companionAppurl: "",
      hasAttachment: false,
    };
    this.handleCheckboxChange = this.handleCheckboxChange.bind(this);
    this.handleComboBoxChange = this.handleComboBoxChange.bind(this);
    this.handleCheckboxFolderChange = this.handleCheckboxFolderChange.bind(this);
  }

  handleCheckboxChange(checked: boolean) {
    this.setState({
      isCheckboxChecked: checked,
    });
  }

  handleCheckboxFolderChange(checked: boolean) {
    this.setState({
      newFolderChecked: checked,
    });
  }

  handleComboBoxChange = (
    _event: React.FormEvent<IComboBox>,
    _option?: IComboBoxOption,
    _index?: number,
    _value?: string,
    id?: string
  ) => {
    this.setState({ selectedSite: _option ? _option.text : null });
    this.setState({ siteValue: id || "" }, async () => {
      try {
        const userAgentApplication = this.context;
        const authScopes = ["https://graph.microsoft.com/Sites.Read.All"];
        const accessToken = await userAgentApplication.acquireTokenSilent({ scopes: authScopes });
        fetch(`https://graph.microsoft.com/v1.0/sites/${this.state.siteValue}/drives`, {
          headers: {
            Authorization: `Bearer ${accessToken.accessToken}`,
          },
        })
          .then((response) => response.json())
          .then(async (data) => {
            // const library = data.value.map((item: any) => ({
            //   id: item.id,
            //   name: item.name,
            //   webUrl: item.webUrl,
            // }));
            // this.setState({
            //   library,
            // });
            let libraryPromises = data.value.map(async (library) => {
              const folderResponse = await fetch(
                `https://graph.microsoft.com/v1.0/sites/${this.state.siteValue}/drives/${library.id}/root/children`,
                {
                  headers: {
                    Authorization: `Bearer ${accessToken.accessToken}`,
                  },
                }
              );
              const folderData = await folderResponse.json();

              // Always include the parent library
              let arr_folders: any[] = [
                {
                  name: library.name,
                  webUrl: library.webUrl,
                  parentId: library.id,
                },
              ];

              // If there are child folders, append them to the array
              if (folderData.value.length > 0) {
                const folders = folderData.value.filter((item: any) => item.folder);
                const childFolders = folders.map((folder: any) => {
                  return {
                    id: folder.id,
                    name: `${library.name} > ${folder.name}`,
                    webUrl: folder.webUrl,
                    parentId: library.id,
                  };
                });
                arr_folders.push(...childFolders);
              }
              return arr_folders;
            });
            let library = await Promise.all(libraryPromises);
            library = library.flat();
            this.setState({
              library: library,
            });
          })
          .catch((error) => console.error("Error:", error));
      } catch (error) {
        console.error("Error while acquiring token or fetching data:", error);
      }
    });
  };

  handleDocumentComboBoxChange = (
    _event: React.FormEvent<IComboBox>,
    item?: IComboBoxOption,
    _index?: number,
    _value?: string,
    id?: string,
    _parentId?: string
  ) => {
    this.setState({ selectedLibrary: item ? item.text : null });
    if (item) {
      this.setState({ libraryId: id || "" }, async () => this.state.libraryId);
      this.setState({ parentLibraryId: _parentId });
    }
  };

  // userAgentApplication = new UserAgentApplication({
  //   auth: {
  //     clientId: "43da5aa8-f8bc-4cb9-83a9-1b2efba5ffb6", // process.env.CLIENT_ID, //
  //     authority: "https://login.microsoftonline.com/e571e05f-df5a-4cac-af8b-272965d6a1cc",
  //   },
  // });

  async componentDidMount() {
    const authScopes = ["https://graph.microsoft.com/Sites.Read.All"];

    while (!this.context) {
      await new Promise((resolve) => setTimeout(resolve, 100)); // Wait until UserAgentApplication is ready
    }

    const userAgentApplication = this.context;

    let accounts = userAgentApplication.getAllAccounts();

    if (accounts.length === 0) {
      // If there are no accounts, then we prompt the user to log in
      try {
        await userAgentApplication.loginPopup({ scopes: authScopes });
      } catch (error) {
        console.error("Login error:", error);
        return;
      }
    }

    let accessToken;
    try {
      // Try to acquire a token
      const response = await userAgentApplication.acquireTokenSilent({
        scopes: authScopes,
        account: userAgentApplication.getAllAccounts()[0],
      });
      accessToken = response.accessToken;
    } catch (error) {
      console.error("Error acquiring token:", error);
      return;
    }

    try {
      // Fetch the root site ID
      const rootSiteResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/root`, {
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      let attach = Office.context.mailbox.item.attachments;
      const realAttachments = attach.filter((attachment) => !attachment.isInline);
      // const hasAttachments = realAttachments.length > 0;
      if (realAttachments.length > 0) {
        this.setState({ hasAttachment: true });
      } else {
        this.setState({ hasAttachment: false });
      }

      if (!rootSiteResponse.ok) {
        throw new Error(`HTTP error! status: ${rootSiteResponse.status}`);
      }

      const rootSiteData = await rootSiteResponse.json();
      const siteId = rootSiteData.siteCollection.hostname;
      let listId;

      // Customize listId based on siteId
      switch (siteId) {
        case "entinvestigations.sharepoint.com":
          listId = "6d6bedc8-5132-49e3-9317-1e19ca352dc5";
          this.setState({
            companionAppurl:
              "https://apps.powerapps.com/play/e/default-6b192610-a943-4065-b690-745cafbc9906/a/9bcc29f6-c807-4e4f-bbab-28b087c9d37a/?hidenavbar=true",
          });
          break;
        case "accountabilityvisibilitycom.sharepoint.com":
          listId = "5c68b26c-3d2e-464c-8ed1-e65fa3f0d50b";
          this.setState({
            companionAppurl:
              "https://apps.powerapps.com/play/e/6ca6301f-828e-e5a7-b452-6b6898ef4ac5/a/1abfb560-7ea0-4f5e-ad5a-4e984b1b4a0f/?hidenavbar=true",
          });
          break;
        default:
          console.error("No matching siteId found!");
          return;
      }

      // Fetch list items with the resolved siteId and listId
      const listItemsResponse = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields`,
        {
          method: "GET",
          headers: {
            Accept: "application/json;odata.metadata=minimal",
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const listItemsData = await listItemsResponse.json();

      const updatedSite = listItemsData.value.map((item) => ({
        title: item.fields.Title,
        url: item.fields.URL.Url,
        siteId: item.fields.SiteID,
      }));

      this.setState({ site: updatedSite });
    } catch (error) {
      console.error("Error:", error);
    }
  }

  sanitizeFilename = (inputName: string) => {
    var illegalRe = /[\/\?<>\\:\*\|":]/g;
    var controlRe = /[\x00-\x1f\x80-\x9f]/g;
    var reservedRe = /^\.+$/;
    var windowsReservedRe = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])(\..*)?$/i;
    var windowsTrailingRe = /[\. ]+$/;

    var sanitized = inputName
      .replace(illegalRe, "")
      .replace(controlRe, "")
      .replace(reservedRe, "")
      .replace(windowsReservedRe, "")
      .replace(windowsTrailingRe, "");

    return sanitized;
  };

  click = async () => {
    const {
      selectedSite,
      selectedLibrary,
      siteValue,
      libraryId,
      parentLibraryId,
      isCheckboxChecked,
      newFolderChecked,
    } = this.state;
    const userAgentApplication = this.context;
    const authScopes = ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Sites.ReadWrite.All"];

    const setLoadingState = (isLoading, showSuccess = false) => {
      this.setState({ isLoading, showSuccess });
      if (showSuccess) {
        setTimeout(() => {
          this.setState({ isLoading: false, showSuccess: false });
        }, 9000);
      }
    };

    if (!selectedSite || !selectedLibrary) {
      this.setState({ showErrors: true });
      return; // stop the execution if a necessary selection is missing
    }
    console.log("parentLibraryId: ", parentLibraryId);
    console.log("libraryId: ", libraryId);
    console.log("newFolderChecked: ", newFolderChecked);
    setLoadingState(true);

    try {
      let accessToken = userAgentApplication.getAccount()
        ? await userAgentApplication.acquireTokenSilent({ scopes: authScopes })
        : await userAgentApplication.loginPopup({ scopes: authScopes });

      let userId = encodeURIComponent(accessToken.account.userName);
      accessToken = accessToken.accessToken;
      let itemId = Office.context.mailbox.item.itemId;
      let messageId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v1_0);
      messageId = encodeURIComponent(messageId);
      let subject = Office.context.mailbox.item.subject;
      let fileName = `${this.sanitizeFilename(subject)}.eml`;
      let folderName = `${this.sanitizeFilename(subject)}`;

      let baseApiUrl = `https://graph.microsoft.com/v1.0/sites/${siteValue}/drives/`;
      let targetFolderId;

      if (newFolderChecked && !libraryId) {
        console.log(1);
        baseApiUrl += `${parentLibraryId}/items`;
      } else if (newFolderChecked && libraryId) {
        console.log(2);
        baseApiUrl += `${parentLibraryId}/items`; ///${libraryId}/children
      } else if (libraryId) {
        console.log(3);
        baseApiUrl += `${parentLibraryId}/items`;
        targetFolderId = `/${libraryId}`;
      } else {
        console.log(4);
        baseApiUrl += `${parentLibraryId}/root`;
        // targetFolderId = `/${parentLibraryId}`;
      }
      console.log("baseApiUrl: ", baseApiUrl);
      let finalPostUrl;
      if (libraryId) {
        finalPostUrl = `${baseApiUrl}/${libraryId}/children`;
      } else {
        finalPostUrl = `${baseApiUrl}`;
      }

      if (newFolderChecked) {
        const response = await fetch(finalPostUrl, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "rename",
          }),
        });

        const data = await response.json();
        targetFolderId = `/${data.id}`;
      }
      console.log("targetFolderId: ", targetFolderId);

      // Upload the email content
      let response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}/$value`, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      let data = await response.text();
      let finalPutUrl;
      if (targetFolderId) {
        finalPutUrl = `${baseApiUrl}${targetFolderId}:/${fileName}:/content`;
      } else {
        finalPutUrl = `${baseApiUrl}:/${fileName}:/content`;
      }

      response = await fetch(finalPutUrl, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "message/rfc822", // This is the MIME type for .eml files
        },
        body: data,
      });

      data = await response.json();
      setLoadingState(false, true);
      console.log("File uploaded successfully. Details:", data);

      if (isCheckboxChecked) {
        response = await fetch(`https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}/attachments`, {
          method: "GET",
          headers: { Authorization: `Bearer ${accessToken}` },
        });

        data = await response.json();
        let attachments = (data as any).value;

        for (let attachment of attachments) {
          if (!attachment.isInline) {
            response = await fetch(
              `https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}/attachments/${attachment.id}/$value`,
              {
                method: "GET",
                headers: { Authorization: `Bearer ${accessToken}` },
              }
            );

            let blob = await response.blob();
            let attachmentFileName = attachment.name;
            let finalPutAttachmentUrl;
            if (targetFolderId) {
              finalPutAttachmentUrl = `${baseApiUrl}${targetFolderId}:/${attachmentFileName}:/content`;
            } else {
              finalPutAttachmentUrl = `${baseApiUrl}:/${attachmentFileName}:/content`;
            }
            response = await fetch(finalPutAttachmentUrl, {
              method: "PUT",
              headers: { Authorization: `Bearer ${accessToken}` },
              body: blob,
            });

            data = await response.json();
            setLoadingState(false, true);
            console.log("Attachment uploaded successfully. Details:", data);
          }
        }
      }
    } catch (error) {
      console.error("Error:", error);
      setLoadingState(false);
    }
  };

  render(): React.JSX.Element {
    const { title, isOfficeInitialized } = this.props;
    const emailAddress = Office.context.mailbox.userProfile.emailAddress;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app bodys."
        />
      );
    }
    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        {this.state.showSuccess && (
          <MessageBar messageBarType={MessageBarType.success}>Uploaded successfully!</MessageBar>
        )}
        <Sites
          showErrors={this.state.showErrors}
          items={this.state.site}
          onComboBoxChange={this.handleComboBoxChange}
          documentItems={this.state.library}
          onDocumentComboBoxChange={this.handleDocumentComboBoxChange}
          companionURL={this.state.companionAppurl}
        ></Sites>
        <HeroList
          message={emailAddress}
          items={this.state.listItems}
          onCheckboxChange={this.handleCheckboxChange}
          onCheckboxFolderChange={this.handleCheckboxFolderChange}
          hasAttachment={this.state.hasAttachment}
        >
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: this.state.isLoading ? "" : "CheckMark" }} // remove the icon when loading
            onClick={this.click}
            disabled={this.state.isLoading}
          >
            {this.state.isLoading ? <Spinner size={SpinnerSize.small} /> : "Submit"}
          </DefaultButton>
        </HeroList>
        <Footer />
      </div>
    );
  }
}
