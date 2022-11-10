import React, { useState, useEffect } from 'react';
import moment from 'moment';
import { CopyToClipboard } from 'react-copy-to-clipboard';
import { AccessToken } from '@azure/identity';
import {
  ShieldKeyholeIcon, UserFriendsIcon, ApprovalsAppbarIcon,
  Button, Flex, Text, Card, CardHeader, CardBody, CardFooter,
} from '@fluentui/react-northstar';

import Requester from '../utils/TeamsDirRequester';
import * as UTILS from '../utils/common';

const Signin: React.FC = () => {
  const [signinStatus, setSigninStatus] = useState<any>(<Text weight="bold" content="Please Signin" success />);
  const [xscToken, setXscToken] = useState<AccessToken>({} as AccessToken);
  const [loading, setLoading] = useState<boolean>(false);
  const [msProfile, setMsProfile] = useState<string>('idle');
  const [graphLoading, setGraphLoading] = useState<boolean>(false);

  useEffect(() => {
    const getXcsToken = async () => {
      const token = await UTILS.getXcsToken();
      if (!token) {
        setXscToken({ token: 'no token', expiresOnTimestamp: 0 });
      } else {
        setXscToken(token);
      }
    };

    getXcsToken();
  }, []);

  const onSignin = async () => {
    const axiosConfig = {
      url: '/token/signin',
    };
    setLoading(true);
    setSigninStatus(null);
    try {
      const response = await Requester(axiosConfig);
      // set localStorage
      localStorage.setItem('uuid', response.data.uuid);
      localStorage.setItem('sessionid', response.data.sessionid);
      setSigninStatus(response.statusText);
    } catch (e) {
      const error = e.response;
      setSigninStatus(<Text important error content={`${error.status} | Sign in ${error.statusText}`} />);
    }
    setLoading(false);
  };

  const getMSProfile = async () => {
    setGraphLoading(true);

    try {
      const res = await UTILS.msGraphMyProfile();
      console.info(res);
      setMsProfile(JSON.stringify(res));
    } catch (e) {
      const error = e.response;
      console.error(error);
      setMsProfile(JSON.stringify(error));
    }
    setGraphLoading(false);
  };

  return (
    <div>
      <Flex padding="padding.medium" column gap="gap.medium">
        <Flex column gap="gap.medium">
          <Card fluid>
            <CardHeader>
              <Flex vAlign="center" gap="gap.small">
                <ShieldKeyholeIcon size="larger" />
                <Text weight="bold" size="larger" content="X-cs token" />
              </Flex>
            </CardHeader>
            <CardBody>
              <Text content={xscToken.token} styles={{ wordWrap: 'break-word' }} />
            </CardBody>
            <CardFooter>
              <CopyToClipboard text={xscToken.token}>
                <Flex hAlign="end">
                  <Button
                    tinted
                    primary
                    content="Copy to clipboard"
                  />
                </Flex>
              </CopyToClipboard>
            </CardFooter>
          </Card>
        </Flex>
        <Flex gap="gap.medium">
          <Card>
            <CardHeader>
              <Flex vAlign="center" gap="gap.small">
                <ApprovalsAppbarIcon />
                <Text weight="bold" size="medium" content="Expired" />
              </Flex>
            </CardHeader>
            <CardBody>
              <Text content={moment(xscToken?.expiresOnTimestamp).format('MMMM Do YYYY, h:mm:ss a')} />
            </CardBody>
          </Card>
          <Card centered>
            <CardHeader>
              <Flex vAlign="center" gap="gap.small">
                <UserFriendsIcon />
                <Text weight="bold" size="large" content="Signin Refinitiv PD" />
              </Flex>
            </CardHeader>
            <CardBody>
              <Flex padding="padding.medium">
                <Button
                  primary
                  loading={loading}
                  content={loading ? 'Loading' : 'Sign in'}
                  onClick={() => onSignin()}
                />
              </Flex>
            </CardBody>
            <CardFooter>
              <Flex hAlign="center">{signinStatus}</Flex>
            </CardFooter>
          </Card>
        </Flex>
        <Flex column gap="gap.medium">
          <Card fluid>
            <CardHeader>
              <Flex vAlign="center" gap="gap.small">
                <ShieldKeyholeIcon size="larger" />
                <Text weight="bold" size="larger" content="get MS graph api" />
              </Flex>
            </CardHeader>
            <CardBody>
              <Text content={msProfile} styles={{ wordWrap: 'break-word' }} />
            </CardBody>
            <CardFooter>
              <Flex hAlign="end">
                <Button
                  onClick={() => getMSProfile()}
                  primary
                  loading={graphLoading}
                  content={graphLoading ? 'Loading' : 'get MS profile'}
                />
              </Flex>
            </CardFooter>
          </Card>
        </Flex>
      </Flex>
    </div>
  );
};

export default Signin;
