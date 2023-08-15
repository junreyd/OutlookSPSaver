/* eslint-disable @typescript-eslint/no-unused-vars */
// import * as React from "react";

// export interface HeroListItem {
//   icon: string;
//   primaryText: string;
// }

// export interface HeroListProps {
//   message: string;
//   items: HeroListItem[];
//   children: any;
//   onCheckboxChange: (checked: boolean) => void;
//   onCheckboxFolderChange: (checked: boolean) => void;
// }

// export interface HeroListState {
//   isCheckboxChecked: boolean;
//   newFolderChecked: boolean;
// }

// export default class HeroList extends React.Component<HeroListProps, HeroListState> {
//   constructor(props: HeroListProps) {
//     super(props);
//     this.state = {
//       isCheckboxChecked: false,
//       newFolderChecked: false,
//     };

//     this.handleCheckboxChange = this.handleCheckboxChange.bind(this);
//     this.handleCheckboxFolderChange = this.handleCheckboxFolderChange.bind(this);
//   }

//   handleCheckboxChange(event: React.ChangeEvent<HTMLInputElement>) {
//     this.props.onCheckboxChange(event.target.checked);
//   }

//   handleCheckboxFolderChange(event: React.ChangeEvent<HTMLInputElement>) {
//     this.props.onCheckboxFolderChange(event.target.checked);
//   }

//   render() {
//     const { children, items, message } = this.props;

//     const listItems = items.map((item, index) => (
//       <li className="ms-ListItem" key={index}>
//         <i className={`ms-Icon ms-Icon--${item.icon}`}></i>
//         <span className="ms-font-m ms-fontColor-neutralPrimary">{item.primaryText}</span>
//       </li>
//     ));
//     return (
//       <main className="ms-welcome__main">
//         <label style={{ marginBottom: "10px" }}>
//           <input type="checkbox" onChange={this.handleCheckboxFolderChange} />
//           Save in new folder
//         </label>
//         <label>
//           <input type="checkbox" onChange={this.handleCheckboxChange} />
//           Include Attachment
//         </label>
//         <br></br>
//         {children}
//       </main>
//     );
//   }
// }

import * as React from "react";
import { Text, Stack, Separator, IStackTokens } from "@fluentui/react";

export interface HeroListItem {
  icon: string;
  primaryText: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
  children: any;
  onCheckboxChange: (checked: boolean) => void;
  onCheckboxFolderChange: (checked: boolean) => void;
  hasAttachment: boolean;
}

export interface HeroListState {
  isCheckboxChecked: boolean;
  newFolderChecked: boolean;
}

export default class HeroList extends React.Component<HeroListProps, HeroListState> {
  constructor(props: HeroListProps) {
    super(props);
    this.state = {
      isCheckboxChecked: false,
      newFolderChecked: false,
    };
    this.handleCheckboxChange = this.handleCheckboxChange.bind(this);
    this.handleCheckboxFolderChange = this.handleCheckboxFolderChange.bind(this);
  }

  handleCheckboxChange(event: React.ChangeEvent<HTMLInputElement>) {
    this.props.onCheckboxChange(event.target.checked);
  }

  handleCheckboxFolderChange(event: React.ChangeEvent<HTMLInputElement>) {
    this.props.onCheckboxFolderChange(event.target.checked);
  }

  render() {
    const { children, items, message } = this.props;
    const listItems = items.map((item, index) => (
      <li className="ms-ListItem" key={index}>
        <i className={`ms-Icon ms-Icon--${item.icon}`}></i>
        <span className="ms-font-m ms-fontColor-neutralPrimary">{item.primaryText}</span>
      </li>
    ));
    const stackTokens: IStackTokens = {
      childrenGap: 10, // Adjust the gap between buttons as needed
    };
    return (
      <main className="ms-welcome__main">
        <Stack horizontal tokens={stackTokens}>
          <label style={{ marginBottom: "10px" }}>
            <input type="checkbox" onChange={this.handleCheckboxFolderChange} />
            Save in folder
          </label>
          {this.props.hasAttachment && (
            <label>
              <input type="checkbox" onChange={this.handleCheckboxChange} />
              Include attachments
            </label>
          )}
        </Stack>
        <br></br>
        {children}
      </main>
    );
  }
}
