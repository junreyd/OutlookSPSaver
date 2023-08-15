/* eslint-disable no-irregular-whitespace */
// import * as React from "react";
// import { Stack, Text } from "@fluentui/react";

// export const Footer: React.FC = () => {
//   return (
//     <Stack
//       horizontalAlign="center"
//       verticalAlign="center"
//       styles={{ root: { position: "fixed", bottom: 0, width: "100%", height: 50, backgroundColor: "#f3f2f1" } }}
//     >
//       <Text variant="small">© 2023 Pacer Solutions. All rights reserved.</Text>
//     </Stack>
//   );
// };

// export default Footer;

import * as React from "react";
import { Stack, Text } from "@fluentui/react";
export const Footer: React.FC = () => {
  return (
    <Stack horizontalAlign="center" verticalAlign="center" className="footer">
      <Text variant="small">
        <span>© 2023 Pacer Solutions. All rights reserved.</span>
      </Text>
      ​
    </Stack>
  );
};
export default Footer;
