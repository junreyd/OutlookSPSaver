/* eslint-disable no-undef */
export class ConfigService {
  static async getClientIdAndAuthority() {
    const response = await fetch("https://outlookspsaver.azurewebsites.net/api/HttpTrigger1");
    const data = await response.json();

    return {
      clientId: data.clientId,
      authority: data.authority,
    };
  }
}
