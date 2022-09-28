import * as React from 'react';
import styles from './ListAPalooza.module.scss';
import { IListAPaloozaProps, IListAPaloozaState } from './IListAPaloozaProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { getSP } from "../pnpjsConfig";
import { SPFI, spfi } from "@pnp/sp"; // eslint-disable-line  @typescript-eslint/no-unused-vars
import { Logger, LogLevel } from "@pnp/logging";

const siteUrl = '/sites/SharePointLists'; // eslint-disable-line  @typescript-eslint/no-unused-vars

export default class ListAPalooza extends React.Component<IListAPaloozaProps, IListAPaloozaState> {
  private LOG_SOURCE = "🅿PnPjsExample";
  private _sp: SPFI; // eslint-disable-line  @typescript-eslint/no-unused-vars

  constructor(props: IListAPaloozaProps) {
    super(props);
    // set initial state
    this.state = {
      items: [],
    };
    this._sp = getSP();
  }

  public async componentDidMount(): Promise<void> {
    // read all file sizes from Documents library
    await this._getRemoteLists();
  }




  public render(): React.ReactElement<IListAPaloozaProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.listAPalooza} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }

  // private _getRemoteLists = async (): Promise<void> => {  //Does not execute anything during 
  private async _getRemoteLists(): Promise<void> {
    console.log('Is this executing????');
    try {


      /**
       * In the past, I would just use this:
       * 
       * try { 
          thisWebInstance = Web(webURL);
          allLists = await thisWebInstance.lists.select('*,HasUniqueRoleAssignments').get();
       * 
       * }
       */
      // get the default document library 'Documents'
      const lists = await this.props.sp.web.lists.select('*')();
      Logger.write(`${this.LOG_SOURCE} (getLists) `, LogLevel.Info);
      console.log('lists', lists );

      this.setState({ items: lists });
    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (getLists) - ${JSON.stringify(err)} - `, LogLevel.Error);
    }
  }

}
