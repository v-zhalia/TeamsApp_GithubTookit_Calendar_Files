import * as React from 'react';
import { Person, MgtTemplateProps} from '@microsoft/mgt-react';
import { makeStyles} from '@fluentui/react-components';
import { updatePresence,clearPreferredPresence,getPresence,Availability, Activity } from '../services/graphPresenceService';
import { PresenceGet } from './PresenceGet';


const useStyles = makeStyles({
  personAvatar: {
    '--person-avatar-size': '48px'
  },
  container: {
    display: 'flex',
    alignItems: 'flex-start',
    flexDirection : 'column',
  },
  statusContainer:{
    display: 'flex',
    width: '100%',
    marginTop: '10px'
  }
});

export const LoginFlyout: React.FunctionComponent<MgtTemplateProps> = (props) => {
  const styles = useStyles();
  const { personDetails } = props.dataContext;
  React.useEffect(() => {
    const loginAndSetPresence = async () => {
      let pre = await getPresence();
        if (pre.availability === Availability.Offline || pre.availability === Availability.Away) {
          await clearPreferredPresence();
          await updatePresence(Availability.Available, Activity.Available);
        }
        else{
        }
    };
    loginAndSetPresence(); // refersh presence on first load
  }, []);

  return (
    <div className={styles.container}>
    <Person 
    userId={personDetails.id} 
    className={styles.personAvatar}
    view={'fourlines'}
    />
    <div className={styles.statusContainer}>
      <PresenceGet container={"flyoutLogin"}></PresenceGet>
    </div>

  </div>
  
  );
};
