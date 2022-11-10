import { AccessToken } from '@azure/identity';
import axios, { AxiosRequestConfig } from 'axios';
import {
  endpointOptionType,
  endpointType,
} from '../components/types/utils.type';

// eslint-disable-next-line import/no-cycle
import { getXcsToken } from './common';

const endpoint = process.env.REACT_APP_TEAMS_DIR_ENDPOINT;
const MsGraphEndpoint = process.env.REACT_APP_TEAMSFX_ENDPOINT;

export default async function request(
  config: AxiosRequestConfig, endpointOption?: endpointOptionType,
) {
  switch (endpointOption) {
    case endpointType.MSGRAPH:
      config.baseURL = MsGraphEndpoint;
      break;
    default:
      config.baseURL = endpoint;
  }

  const xcsToken: AccessToken | null = await getXcsToken();
  config.headers = {
    ...config.headers,
    'x-cs-token': xcsToken?.token,
  };
  return axios.request(config);
}
