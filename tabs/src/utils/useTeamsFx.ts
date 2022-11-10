import {
  loadConfiguration, ResourceType, LogLevel, setLogLevel, setLogFunction,
} from '@microsoft/teamsfx';
import { useTeams } from 'msteams-react-base-component';
import { useData } from './useData';

const teamsfxEndpoint = process.env.REACT_APP_TEAMSFX_ENDPOINT;
const startLoginPageUrl = process.env.REACT_APP_START_LOGIN_PAGE_URL;
const functionEndpoint = process.env.REACT_APP_FUNC_ENDPOINT;
const clientId = process.env.REACT_APP_CLIENT_ID;

// TODO fix this when the SDK stops hiding global state!
let initialized = false;

// eslint-disable-next-line import/prefer-default-export
export function useTeamsFx() {
  const [result] = useTeams({});
  const { error, loading } = useData(async () => {
    if (!initialized) {
      if (process.env.NODE_ENV === 'development') {
        setLogLevel(LogLevel.Verbose);
        setLogFunction((leve: LogLevel, message: string) => { console.info(message); });
      }
      loadConfiguration({
        authentication: {
          initiateLoginEndpoint: startLoginPageUrl,
          simpleAuthEndpoint: teamsfxEndpoint,
          clientId,
        },
        resources: [
          {
            type: ResourceType.API,
            name: 'default',
            properties: {
              endpoint: functionEndpoint,
            },
          },
        ],
      });
      initialized = true;
    }
  });
  const isInTeams = true;
  return {
    error, loading, isInTeams, ...result,
  };
}
