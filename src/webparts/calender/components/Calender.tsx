import { Text } from 'office-ui-fabric-react';
import * as React from 'react';
import styles from './Calender.module.scss';
import { ICalenderProps } from './ICalenderProps';

export default class Calender extends React.Component<ICalenderProps, {}> {
  public render(): React.ReactElement<ICalenderProps> {
    const {
    } = this.props;

    return (
      <section className={`${styles.calender}`}>
        <Text>Hello Calender WP</Text>
      </section>
    );
  }
}
