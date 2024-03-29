import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { AppComponent } from './app.component';
import { MsalModule, MsalService, MsalGuard, MsalInterceptor, MsalBroadcastService, MsalRedirectComponent } from "@azure/msal-angular";
import { PublicClientApplication, InteractionType, BrowserCacheLocation } from "@azure/msal-browser";

import * as config from '../config'

@NgModule({
    declarations: [
        AppComponent
    ],
    imports: [
        MsalModule.forRoot( new PublicClientApplication({ // MSAL Configuration
            auth: {
                clientId: config.clientId,
                authority: config.authorityUrl,
                redirectUri: "http://localhost:4200",
            },
            cache: {
                cacheLocation : BrowserCacheLocation.LocalStorage,
                storeAuthStateInCookie: true, // set to true for IE 11
            },
            system: {
                loggerOptions: {
                    loggerCallback: () => {},
                    piiLoggingEnabled: false
                }
            }
        }), {
            interactionType: InteractionType.Popup,
            authRequest: {
                scopes:['https://analysis.windows.net/powerbi/api/Report.Read.All']
            } // MSAL Guard Configuration
        }, {
            interactionType: InteractionType.Popup, // MSAL Interceptor Configuration
            protectedResourceMap: new Map([
                ['https://analysis.windows.net/powerbi/api/Report.Read.All', ['Report.Read.All']],
            ])}),
        BrowserModule,
        HttpClientModule
    ],
    providers: [
        {
            provide: HTTP_INTERCEPTORS,
            useClass: MsalInterceptor,
            multi: true
        },
        MsalService,
        MsalGuard,
        MsalBroadcastService,
    ],
    bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule {}