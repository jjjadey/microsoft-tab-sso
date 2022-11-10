import { TeamsUserCredential } from '@microsoft/teamsfx';
import { AccessToken } from '@azure/identity';
import { SizeValue } from '@fluentui/react-northstar';
import { AxiosRequestConfig } from 'axios';
import * as microsoftTeams from '@microsoft/teams-js';

// function
// eslint-disable-next-line import/no-cycle
import Requester from './TeamsDirRequester';

// eslint-disable-next-line consistent-return
export const msGraphApi = async (axiosConfig: AxiosRequestConfig) => {
  try {
    return (await Requester(axiosConfig, 'msGraph')).data;
  } catch (error) {
    // client display a pop-up auth box
    console.error(error);
  }
};

// Get the user profile from our web service
export const msGraphMyProfile = async () => {
  const config: AxiosRequestConfig = {
    url: '/userProfile',
    method: 'get',
    headers: {
      'Content-Type': 'application/json',
    },
  };
  return msGraphApi(config);
};

// Display the consent pop-up if needed
export const showConsentPopup = async () => {
  let scope = process.env.REACT_APP_SCOPE?.split(',');
  scope = scope?.map((o) => `https://graph.microsoft.com/${o}`);

  let authUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL || 'https://localhost:53000/auth-start.html';

  authUrl += `?clientId=${process.env.REACT_APP_CLIENT_ID || ''}`;

  authUrl += `&scope=${encodeURIComponent(scope?.join(' ') || '')}`;

  microsoftTeams.authentication.authenticate({
    //   url: window.location.origin + "/auth-start.html",
    url: authUrl,
    width: 600,
    height: 535,
    successCallback: (result) => {
      console.info('Got success callback');
      console.info(result);
      // window.location.reload();
    },
    failureCallback: (err) => {
      console.error(err);
    },
  });
};

export const getMsPhoto = async (msUuids: string[], size?: SizeValue) => {
  const config: AxiosRequestConfig = {
    url: '/getMsPhotos',
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
    },
    data: {
      msUuids,
      size,
    },
  };
  const res = msGraphApi(config);
  return res;
};

export const getXcsToken = async () => {
  const credential = new TeamsUserCredential();
  const token: AccessToken | null = await credential.getToken('');
  return token;
};
