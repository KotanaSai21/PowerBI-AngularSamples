
import { Component, OnInit, Inject, OnDestroy, ViewChild, ElementRef } from '@angular/core';
import { MsalService, MsalBroadcastService, MSAL_GUARD_CONFIG, MsalGuardConfiguration } from '@azure/msal-angular';
import { AuthenticationResult, InteractionStatus, PopupRequest, RedirectRequest, EventMessage, EventType } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

import { service, factories, models, IEmbedConfiguration } from "powerbi-client";
import * as config from '../config';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit, OnDestroy {
  title = 'Angular 16 - MSAL Angular v3 Sample';
  isIframe = false;
  loginDisplay = false;
  @ViewChild('embedContainer') private reportContainer!: ElementRef<HTMLDivElement>;

  displayMessage = "Displaymessage";
  powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);
  accessToken: string = "";
  embedUrl = "";
  mock:string | undefined = "hy "
  private readonly _destroying$ = new Subject<void>();

  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private authService: MsalService,
    private msalBroadcastService: MsalBroadcastService
  ) {
    
  }

  ngOnInit(): void {
    this.isIframe = window !== window.parent && !window.opener; // Remove this line to use Angular Universal
    this.setLoginDisplay();

    this.authService.instance.enableAccountStorageEvents(); // Optional - This will enable ACCOUNT_ADDED and ACCOUNT_REMOVED events emitted when a user logs in or out of another tab or window
    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.ACCOUNT_ADDED || msg.eventType === EventType.ACCOUNT_REMOVED),
      )
      .subscribe((result: EventMessage) => {
        if (this.authService.instance.getAllAccounts().length === 0) {
          window.location.pathname = "/";
        } else {
          this.setLoginDisplay();
        }
      });
    
    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe(() => {
        this.setLoginDisplay();
        this.checkAndSetActiveAccount();
      })
  }

  
  ngAfterViewInit() {
    if(this.accessToken !== "" && this.embedUrl !== "") {
      const embedConfiguration: IEmbedConfiguration = {
        type: "report",
        tokenType: models.TokenType.Embed,
        accessToken: this.accessToken,
        embedUrl: this.embedUrl,
        id: config.reportId,
        /*
        // Enable this setting to remove gray shoulders from embedded report
        settings: {
            background: models.BackgroundType.Transparent
        }
        */
      };

      const report = this.powerbi.embed(this.reportContainer.nativeElement, embedConfiguration);

      // Clear any other loaded handler events
      report.off("loaded");

      // Triggers when a content schema is successfully loaded
      report.on("loaded", function () {
        console.log("Report load successful");
      });

      // Clear any other rendered handler events
      report.off("rendered");

      // Triggers when a content is successfully embedded in UI
      report.on("rendered", function () {
        console.log("Report render successful");
      });

      // Clear any other error handler event
      report.off("error");

      // Below patch of code is for handling errors that occur during embedding
      report.on("error", function (event) {
        const errorMsg = event.detail;

        // Use errorMsg variable to log error in any destination of choice
        console.error(errorMsg);
      });
    }

    // User input - null check
    else if (config.workspaceId === "" || config.reportId === "") {
      this.displayMessage =  "Please assign values to workspace Id and report Id in Config.ts file";
      return;
    }

    else {
      this.authenticate();
    }
  }

  async authenticate() {
    // console.log("authenticatrrrre");
  }


  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
  }

  checkAndSetActiveAccount(){
    /**
     * If no active account set but there are accounts signed in, sets first account to active account
     * To use active account set here, subscribe to inProgress$ first in your component
     * Note: Basic usage demonstrated. Your app may require more complicated account selection logic
     */
    let activeAccount = this.authService.instance.getActiveAccount();

    if (!activeAccount && this.authService.instance.getAllAccounts().length > 0) {
      let accounts = this.authService.instance.getAllAccounts();
      this.authService.instance.setActiveAccount(accounts[0]);
    }
  }

  loginPopup() {
    if (this.msalGuardConfig.authRequest){
      this.authService.loginPopup({...this.msalGuardConfig.authRequest} as PopupRequest)
        .subscribe((response: AuthenticationResult) => {
          this.authService.instance.setActiveAccount(response.account);
          this.accessToken = response.accessToken;
          this.mock = response.account?.name;

        });
      } else {
        this.authService.loginPopup()
          .subscribe((response: AuthenticationResult) => {
            this.authService.instance.setActiveAccount(response.account);
            this.accessToken = response.accessToken;
            this.mock = response.account?.name;
      });
    }
  }

  logout(popup?: boolean) {
    if (popup) {
      this.authService.logoutPopup({
        mainWindowRedirectUri: "/"
      });
    } else {
      this.authService.logoutRedirect();
    }
  }

  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
  }
}