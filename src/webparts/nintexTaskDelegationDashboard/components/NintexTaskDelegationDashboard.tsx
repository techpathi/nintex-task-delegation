import * as React from 'react';
import styles from './NintexTaskDelegationDashboard.module.scss';
import type { INintexTaskDelegationDashboardProps } from './INintexTaskDelegationDashboardProps';
import { DelegationRunner } from './DelegationRunner/DelegationRunner';

export default function NintexTaskDelegationDashboard(props: INintexTaskDelegationDashboardProps) {
  return (
    <section className={`${styles.nintexTaskDelegationDashboard} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <DelegationRunner
        context={props.context}
        nintexApiBaseUrl={props.nintexApiBaseUrl}
        tokenListUrl={props.tokenListUrl}
        tokenColumnName={props.tokenColumnName}
        tokenTitleValue={props.tokenTitleValue}
        recentDaysLimit={props.recentDaysLimit}
        bufferMinutes={props.bufferMinutes}
      />

    </section>
  );
}
