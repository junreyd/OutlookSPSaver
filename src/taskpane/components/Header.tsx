/* eslint-disable no-undef */
/* eslint-disable @typescript-eslint/no-unused-vars */
// import * as React from "react";
// import { Text, Stack, Separator } from "@fluentui/react";
// import { FontIcon } from "@fluentui/react/lib/Icon";

// export interface HeaderProps {
//   title: string;
//   logo: string;
//   message: string;
// }

// export default class Header extends React.Component<HeaderProps> {
//   render() {
//     const { title, logo, message } = this.props;
//     return (
//       <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
//         <Stack horizontal horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 10 }}>
//           <FontIcon iconName="Upload" className="ms-IconExample" />
//           <Text variant="xLarge">Your Essential Tool for Email &amp; Attachment Archiving to SharePoint</Text>
//         </Stack>
//         <Separator />
//       </section>
//     );
//   }
// }

import * as React from "react";
import { Text, Stack, Separator, IStackTokens, DefaultButton } from "@fluentui/react";
import { FontIcon, Icon } from "@fluentui/react/lib/Icon";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

export default class Header extends React.Component<HeaderProps> {
  handleReload = () => {
    window.location.reload();
  };

  render() {
    const { title, logo, message } = this.props;
    const iconStyles = { marginRight: "8px" };
    const stackTokens: IStackTokens = {
      childrenGap: 10, // Adjust the gap between buttons as needed
    };

    return (
      <Stack horizontal horizontalAlign="space-between" tokens={stackTokens} className="tab-menu">
        <Stack horizontal tokens={stackTokens}>
          <div className="tab-button">
            <Icon iconName="SharepointAppIcon16" style={iconStyles} />
            <span>Save to SharePoint</span>
          </div>
        </Stack>

        <div className="reload-icon-container" onClick={this.handleReload} title="Reload">
          <Icon iconName="Refresh" className="reload-icon" />
        </div>
      </Stack>
    );
  }
}
