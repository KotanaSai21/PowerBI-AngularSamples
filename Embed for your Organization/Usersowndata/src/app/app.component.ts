
import { Component, Inject, ViewChild, ElementRef } from '@angular/core';
import { HttpClient, HttpErrorResponse, HttpHeaders } from '@angular/common/http';
import { MsalService, MSAL_GUARD_CONFIG, MsalGuardConfiguration } from '@azure/msal-angular';
import { AuthenticationResult, PopupRequest } from '@azure/msal-browser';
import { service, factories, models, IEmbedConfiguration } from "powerbi-client";
import { lastValueFrom, of } from 'rxjs';
import { catchError, map} from 'rxjs/operators';

import * as config from '../config';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})

export class AppComponent  {
  @ViewChild('embedContainer') private reportContainer!: ElementRef<HTMLDivElement>;

  loginDisplay = false;
  displayMessage = "Displaymessage";
  accessToken: string = "";
  embedUrl = "";
  accountName: string | undefined = ""
  powerbi = new service.Service(factories.hpmFactory, factories.wpmpFactory, factories.routerFactory);

  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private authService: MsalService,
    private http: HttpClient
  ) {}

  ngAfterViewInit() {
    if (this.accessToken !== "" && this.embedUrl !== "") {
      // Embed the Report
      this.embedReport();
    }

    // User input - null check
    else if (config.workspaceId === "" || config.reportId === "") {
      this.displayMessage = "Please assign values to workspace Id and report Id in Config.ts file";
      return;
    }

    else {
      // this.loginPopup();
    }
  }

  loginPopup() {
    if (this.msalGuardConfig.authRequest) {
      this.authService.loginPopup({ ...this.msalGuardConfig.authRequest } as PopupRequest)
        .subscribe((response: AuthenticationResult) => {
          this.authService.instance.setActiveAccount(response.account);
          this.accessToken = response.accessToken;
          this.accountName = response.account?.name;
        });
    } else {
      this.authService.loginPopup()
        .subscribe((response: AuthenticationResult) => {
          this.authService.instance.setActiveAccount(response.account);
          this.accessToken = response.accessToken;
          this.accountName = response.account?.name;
        });
    }
    this.setLoginDisplay();
  }

  logout(popup?: boolean) {
    if (popup) {
      this.authService.logoutPopup({
        mainWindowRedirectUri: "/"
      });
    } else {
      this.authService.logoutRedirect();
    }
    this.loginDisplay = false;
  }

  // Get the Embed URL for the report
  async getEmbedUrl(): Promise<string>{
    console.log("getEmbedUrl")
    const url = `${config.powerBiApiUrl}v1.0/myorg/groups/${config.workspaceId}/reports/${config.reportId}`;
    const headers = new HttpHeaders({
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + this.accessToken
    });

    const response = this.http.get(url, { headers }).pipe(
      map((response: any) => response.embedUrl),
      catchError((error: HttpErrorResponse) => {
        console.error('An error occurred:', error.error.message);
        return of(error.error.message);
      })
    );

    return lastValueFrom(response);
  }

  // Embeds a Power BI Report
  async embedReport() {
    this.embedUrl =  await this.getEmbedUrl();
    const embedConfiguration: IEmbedConfiguration = {
      type: "report",
      tokenType: models.TokenType.Aad,
      accessToken: this.accessToken,
      embedUrl: this.embedUrl,
      /*
      // Enable this setting to remove gray shoulders from embedded report
      settings: {
          background: models.BackgroundType.Transparent
      }
      */
    };
    console.log("embedReport",this.embedUrl)
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

  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
    console.log(this.loginDisplay)
  }

  ngOnDestroy(): void {
    this.powerbi.reset(this.reportContainer.nativeElement);
  }
}