import * as React from 'react';
import { MgtTemplateProps} from '@microsoft/mgt-react';
import { PresenceGet} from './PresenceGet';

export const SimpleLogin: React.FunctionComponent<MgtTemplateProps> = () => {
    return (
      <div>
      <PresenceGet container = {"simpleLogin"}></PresenceGet>
      </div>
    ); 
}
