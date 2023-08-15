import React, { createContext, useState, useEffect } from "react";
import { UserAgentApplication, Configuration } from "msal";
import { ConfigService } from "./configService";

export const AuthContext = createContext<UserAgentApplication | null>(null);

interface AuthProviderProps {
  children: React.ReactNode;
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ children }) => {
  const [userAgentApplication, setUserAgentApplication] = useState<UserAgentApplication | null>(null);

  useEffect(() => {
    async function initializeUserAgentApplication() {
      const { clientId, authority } = await ConfigService.getClientIdAndAuthority();

      const msalConfig: Configuration = {
        auth: {
          clientId: clientId,
          authority: authority,
        },
      };

      setUserAgentApplication(new UserAgentApplication(msalConfig));
    }

    initializeUserAgentApplication();
  }, []);

  return <AuthContext.Provider value={userAgentApplication}>{children}</AuthContext.Provider>;
};
