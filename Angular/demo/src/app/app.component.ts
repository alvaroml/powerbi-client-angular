// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Component, ElementRef, ViewChild } from '@angular/core';
import { IHttpPostMessageResponse } from 'http-post-message';
import { IReportEmbedConfiguration, models, Page, Report, service, VisualDescriptor } from 'powerbi-client';
import { PowerBIReportEmbedComponent } from 'powerbi-client-angular';
import 'powerbi-report-authoring';
import { errorClass, errorElement, hidden, position, reportUrl, successClass, successElement } from '../constants';
import { HttpService } from './services/http.service';

// Handles the embed config response for embedding
export interface ConfigResponse {
  Id: string;
  EmbedUrl: string;
  EmbedToken: {
    Token: string;
  };
}

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  // Wrapper object to access report properties
  @ViewChild(PowerBIReportEmbedComponent) reportObj!: PowerBIReportEmbedComponent;

  // Div object to show status of the demo app
  @ViewChild('status') private statusRef!: ElementRef<HTMLDivElement>;

  // Embed Report button element of the demo app
  @ViewChild('embedReportBtn') private embedBtnRef!: ElementRef<HTMLButtonElement>;

  // Track Report embedding status
  isEmbedded = false;

  // Overall status message of embedding
  displayMessage = 'The report is bootstrapped. Click Embed Report button to set the access token.';

  // CSS Class to be passed to the wrapper
  // Hide the report container initially
  reportClass = 'report-container hidden';

  // Flag which specify the type of embedding
  phasedEmbeddingFlag = false;

  // Pass the basic embed configurations to the wrapper to bootstrap the report on first load
  // Values for properties like embedUrl, accessToken and settings will be set on click of button
  reportConfig: IReportEmbedConfiguration = {
    type: 'report',
    embedUrl: undefined,
    tokenType: models.TokenType.Embed,
    accessToken: undefined,
    settings: undefined,
  };

  /**
   * Map of event handlers to be applied to the embedded report
   */
  // Update event handlers for the report by redefining the map using this.eventHandlersMap
  // Set event handler to null if event needs to be removed
  // More events can be provided from here
  // https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/handle-events#report-events
  eventHandlersMap = new Map<string, (event?: service.ICustomEvent<any>) => void>([
    ['loaded', () => {
      console.log('Report has loaded');
      // console.log(this.reportConfig);
    }],
    [
      'rendered',
      () => {
        console.log('Report has rendered');

        // Set displayMessage to empty when rendered for the first time
        if (!this.isEmbedded) {
          this.displayMessage = 'Use the buttons above to interact with the report using Power BI Client APIs.';
        }

        // Update embed status
        this.isEmbedded = true;
      },
    ],
    [
      'error',
      (event?: service.ICustomEvent<any>) => {
        if (event) {
          console.error(event.detail);
        }
      },
    ],
    ['visualClicked', () => console.log('visual clicked')],
    ['pageChanged', (event) => console.log(event)],
  ]);

  constructor(public httpService: HttpService, private element: ElementRef<HTMLDivElement>) {}

  /**
   * Embeds report
   *
   * @returns Promise<void>
   */
  async embedReport(): Promise<void> {
    let reportConfigResponse: ConfigResponse;

    // Get the embed config from the service and set the reportConfigResponse
    /*try {
      reportConfigResponse = await this.httpService.getEmbedConfig(reportUrl).toPromise();
      console.log(reportConfigResponse);
    } catch (error: any) {
      // Prepare status message for Embed failure
      await this.prepareDisplayMessageForEmbed(errorElement, errorClass);
      this.displayMessage = `Failed to fetch config for report. Status: ${error.statusText} Status Code: ${error.status}`;
      console.error(this.displayMessage);
      return;
    }*/

    // Update the reportConfig to embed the PowerBI report
    this.reportConfig = {
      ...this.reportConfig,
      id: '1632b01a-712a-436a-b03d-8d99c6cdabbd',
      embedUrl: 'https://app.powerbi.com/reportEmbed?reportId=1632b01a-712a-436a-b03d-8d99c6cdabbd&groupId=ab2b1946-974b-441d-9bca-a5c449ee8415&w=2&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVTLU5PUlRILUNFTlRSQUwtcmVkaXJlY3QuYW5hbHlzaXMud2luZG93cy5uZXQiLCJlbWJlZEZlYXR1cmVzIjp7Im1vZGVybkVtYmVkIjp0cnVlLCJhbmd1bGFyT25seVJlcG9ydEVtYmVkIjp0cnVlLCJjZXJ0aWZpZWRUZWxlbWV0cnlFbWJlZCI6dHJ1ZSwidXNhZ2VNZXRyaWNzVk5leHQiOnRydWUsInNraXBab25lUGF0Y2giOnRydWV9fQ%3d%3d',
      accessToken: 'H4sIAAAAAAAEAC2Xxw6sWBJE_-VtGQlTUMBIvcB779nhvfe05t-n1Op96ioVN-Jk5t9_zOTppyT_898_z77FBGhrdE_UyfyVSFspxuPFQ9Hj-GsUlDCK61E9pEJhr1nC5ORzS1e0UhnUaP7RIiaU7YlE-khWNNNk6vTSc8ghiTlWABKehzRezjlWF5q2DtFY1EiKXwPO36K33LxEgrHCOjhmIim4pd--I3BTWfDFadZ73GyZaC2pYhX9W78dNakD5rb-PFTuR4zfoIHKPiJyOgLrDJN8jSIuPUjmQWVt3aXrTwFtvITW8aEfnmDCADybbyij2Jm-pIVfzlR57J12KtPBkBWW_iIY17tWAg3ML20tDjDWkvh8mF2ZXQFVsb52lQsmp8RTK4EDcrUqjKpobGB0xy15yLiTgs8xKxcduIoWJGYfAHqS6dX328bdnliM2DcD85zDkaJc7rURdzGl_OQxQwT39nU3vDIfqmYY-rHIE0QmuneRGDWa7AEC6VprsIfSU5MCN3ctIioGMcgG8KvNVK3y-JJm5mPKqIpfzMLdmc4A7xilBJNs3sdVbUgkt4VDdVgcQ1eUlZDWnqkO5mA0jky-ANZbW8pDXEp6TZuwKetT2hOF02yhDr-vHqNlUtS5HqZtFtEFKSONvncvnI_dssIohbzG6pMnrTtAlU-kMQmAHazgk761BMuXZ34kbaQxaD6XNiD0g1L2XQbvZPn0PKuOvT2hSvIhqCcqoeGL0ovsxS9zYqdISooxtEOoOSB3I1GH2CqVhENZKjahiA6Mx8s4EMVrX2tAKFsZXna12qwgwhpXyzk20DOHRkY-KxSqqEGxNcnGOwORKahvCzrZgkpkEoUTBcqKE-q-jt37a3cvjj0TA2KWcbl0856fvq6_-Bm4MAA1QY17BeacO-KgiC5cXS_KDIkA6A0IOhxKJOBhPIM96pQhUWLL-g4dVIklCDacN0fgIMjt5OwqBtxXQ--ZbBbVR7m6VoDnCcEhZAx_7KA03Wdo7jI8EtMc725BYbs1e6rqugozO6r1w-VGiEAn7umnCorxialBPbObax4FzYgPigMlo8cdS8WC6N3L9LHUsPweBJINd26i45vRMIQ60GeH5mgj5PtQ6gUYmZC7g6Zfz0KJn8UvHovSS2myK0Lc4k4eILG6F6G_VXZl_UJ0xLRw6-r2nOQxlN3dmNXWRp0-HY2DKOPFvwIfAERowirDdG3uqN63kzOHqV-IlRRYlISTEljGmtDFZxUh1ITCbn5EugQzc6jawJ11jKGvELw8k7PHfOevlfVtVDarCtaS12b6Fap3olA6vUmNykqtBKg81bW7gVbrTOj3WN7-kRpOir6yseZyoQGBCM7KbJYdRcDxma-cK4HmST6HygAp2eiyXu_LEO5CYCrA-C6jtH5fZ5hYjSfhEC4ale_mQITBQmCx5aXC1gk4Z3B5-sacYoL2iuZj-QePfKwUZz_l2ubxHhaBHGrpq1WV5OVehOXfZ0nl7soixjkRzBWv0z1a2B0GI3nwxhaF91ikn3n2ZxK5JrbVwwSW08q_2fgx1kAUj9_LHPJpMSLd394a3R30dunEOHonWyDVkHT10GS9lUkGkttpLO-DB5NpGh9Fg2DxYs4gaLWqDhNpwWyaOLeuxWaPTQ24FGe-6VbVvxMksgL9Kmtz_SE7-w6anNuiPMgppi9SYSSjgYxZLHSCLeKG33pbhO1QZ9ekSfOz5jLfMcTzqNnEsz8vlfh8QOZdwZ4ZbA3xCwZMvV5KXRrO4LUgWeR5_DGiyNdgVRsUIAPR60bjlUyE2vvyt4jsN5PsnzVyYevTAqsxDP0UEGSuqLk_izF6w9fi-z27xt_I-kDaaXik7_bR1YFNdSvKi0op3Pa2b91jp-FRF7xI-QE9JE-6xAeNSpKYJdTjjVI8Csx0HNssSOomS6ajwm-1ojpVNKN53bZZuljnX5j0QRMbXhsX_3PiFSiQ4iwPXFTD264Y-OHIlVWeUhTb4hioJWWip9gS4xAXJaKGSuhavtfTO0MQDvWarhY8HPlo2l4JAUTeScCRO5-iYrOtQizLES87-Kk1OE0z6_6gCgTBR6QklDkNQJtymHhqQV9EqDIV7NzslNKBdpvkyo1NHjiOmNbR6bXu6bK3PY4ye5R14DMmWc1zdA1hH1UJ5ynWfP_AOgCtyphi-YBKhwmuNK3sM8tFvxzS57x-VAP8SYxZQnLuK6JrJitt8eGFKeQ9qUnuDsYjwngNAs7It5Ft5hechcOSHjawybWaYefIe5Cb-NqNEcfhDDClUzsRYtACepWaaVQcXjzB655638RuLjNneQLQpu2OZNRx6xRvw2zOL8zZOcsPOTVSBKOv3IR21Y8QsA8QNjj5KsiyeT8_g1qoAP3WteATy4yS3xddXMUnF4cqYiXK7QEhUC7gQQqeZSOdS4Kmu046WIIjmafJXrYbkxwVUaP0CrFL82lpoxkgxmz6MdXyzwdlQD1X6-XWUXLZmsZU4aEsyrga9GvM5nYGFugpbSSegGC792pjek3L3VjO65pDPc3qAgPH7oxglZq29TLl9nZ1tENFYFcdKF_Qj4vPIZKZGm--p0tZQJS8o2YcL0F0QatcWdZkntLjpa695g_3zb6H6kLvjv8csDheBnYuxB8iqZSHYaQWh7yRa4oH2bNKucDZKj4du05NTS3xDnJ80kA5jouH9V3QCsRXJGNyBu91X6j90xDcJNsJGmuXmKf--uvPf_4w6zPvk1I8vzUz8JUfIeZmUz2O9RNDmBbNfnZzYtO8oqcrkntaXF__0UrpMJvFShgqCkIBJcvGyb5Xj5zZl1l8pJ-N1ZIJRPyc2jLJSG-4QVJ1aJrSJ95O-bmU8ol5WTauqyK9DgxND2pEdQweIFD8IkR0zM-S4jU_v5XVqfkezkrGY-uiWojL7fRbkKzph3PWbPT8MRUeA_pJt_r7QCpXUfWuEZEtM4562vrlGMDl0jUVQrsgYpGJgAX2XMt9Qtw4fvryfds1DUi4R-CXTj0Sw16EG3OvzA71RzgJst0nNh0A9nPVYoYYzfoMci80F5zwIEaGP-DWTmwsLHZAP3KlkU4HKoiV_aYD8v4r8zPXxSr5P5X1OYYcYU6-CwCwOaMC-G-8Vf9UOU01JvuxFr8yhx5NOskAuW421iYuPAzi3WkDiRo2WGciUxm1efUtU-6zjYCfcbKwKrCYMyR7FaTcINigotVhyYh8ghnFxqFK7pryJtCEd4kVCrkiXK2nh3Or6cMxnizi0luchzvcSWh95tANJQ59X-zwubiZSGJgtZbTJhOA2Yk5ANmcst850JvmAWjIWo4jc31rbPNC70K5wbuyR5TjBHZJg4EY3Qa8VrTkHKq9BhOKF_8J_a5rUWWd33QkTt1PM3HfDzS5R8Uyuz99aS2mCUo3N6S_473ZlvPYiO_Q3snH1NkEG38Q0d5S_p4lusupwN1jhClVWJDkzRM9jHhKDbrZDkIKkxZGqW7XT-b__R_bEKSLWg0AAA==.eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVNPVVRILUNFTlRSQUwtVVMtcmVkaXJlY3QuYW5hbHlzaXMud2luZG93cy5uZXQiLCJlbWJlZEZlYXR1cmVzIjp7Im1vZGVybkVtYmVkIjpmYWxzZX19',
    };

    // Get the reference of the report-container div
    const reportDiv = this.element.nativeElement.querySelector('.report-container');
    if (reportDiv) {
      // When Embed report is clicked, show the report container div
      reportDiv.classList.remove(hidden);
    }

    // Get the reference of the display-message div
    const displayMessage = this.element.nativeElement.querySelector('.display-message');
    if (displayMessage) {
      // When Embed report is clicked, change the position of the display-message
      displayMessage.classList.remove(position);
    }

    // Prepare status message for Embed success
    await this.prepareDisplayMessageForEmbed(successElement, successClass);

    // Update the display message
    this.displayMessage = 'Access token is successfully set. Loading Power BI report.';
  }

  /**
   * Handle Report embedding flow
   * @param img Image to show with the display message
   * @param type Type of the message
   *
   * @returns Promise<void>
   */
  async prepareDisplayMessageForEmbed(img: HTMLImageElement, type: string): Promise<void> {
    // Remove the Embed Report button from UI
    this.embedBtnRef.nativeElement.remove();

    // Prepend the Image element to the display message
    this.statusRef.nativeElement.prepend(img);

    // Set type of the message
    this.statusRef.nativeElement.classList.add(type);
  }

  /**
   * Delete visual
   *
   * @returns Promise<void>
   */
  async deleteVisual(): Promise<void> {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error
      this.prepareStatusMessage(errorElement, errorClass);
      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }

    // Get all the pages of the report
    const pages: Page[] = await report.getPages();

    // Check if all the pages of the report deleted
    if (pages.length === 0) {
      // Prepare status message for Error
      this.prepareStatusMessage(errorElement, errorClass);
      this.displayMessage = 'No pages found.';
      console.log(this.displayMessage);
      return;
    }

    // Get active page of the report
    const activePage: Page | undefined = pages.find((page) => page.isActive);

    if (activePage) {
      // Get all visuals in the active page of the report
      const visuals: VisualDescriptor[] = await activePage.getVisuals();

      if (visuals.length === 0) {
        // Prepare status message for Error
        this.prepareStatusMessage(errorElement, errorClass);
        this.displayMessage = 'No visuals found.';
        console.log(this.displayMessage);
        return;
      }

      // Get first visible visual
      const visual: VisualDescriptor | undefined = visuals.find((v) => v.layout.displayState?.mode === models.VisualContainerDisplayMode.Visible);

      // No visible visual found
      if (!visual) {
        // Prepare status message for Error
        this.prepareStatusMessage(errorElement, errorClass);
        this.displayMessage = 'No visible visual available to delete.';
        console.log(this.displayMessage);
        return;
      }

      try {
        // Delete the visual using powerbi-report-authoring
        // For more information: https://docs.microsoft.com/en-us/javascript/api/overview/powerbi/report-authoring-overview
        const response = await activePage.deleteVisual(visual.name);

        // Prepare status message for success
        this.prepareStatusMessage(successElement, successClass);
        this.displayMessage = `${visual.type} visual was deleted.`;
        console.log(this.displayMessage);

        return response;
      } catch (error) {
        console.error(error);
      }
    }
  }

  /**
   * Hide Filter Pane
   *
   * @returns Promise<IHttpPostMessageResponse<void> | undefined>
   */
  async hideFilterPane(): Promise<IHttpPostMessageResponse<void> | undefined> {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();

    if (!report) {
      // Prepare status message for Error
      this.prepareStatusMessage(errorElement, errorClass);
      this.displayMessage = 'Report not available.';
      console.log(this.displayMessage);
      return;
    }

    // New settings to hide filter pane
    const settings = {
      panes: {
        filters: {
          expanded: false,
          visible: false,
        },
      },
    };

    try {
      const response = await report.updateSettings(settings);

      // Prepare status message for success
      this.prepareStatusMessage(successElement, successClass);
      this.displayMessage = 'Filter pane is hidden.';
      console.log(this.displayMessage);

      return response;
    } catch (error) {
      console.error(error);
      return;
    }
  }

  /**
   * Set data selected event
   *
   * @returns void
   */
  setDataSelectedEvent(): void {
    // Adding dataSelected event in eventHandlersMap
    this.eventHandlersMap = new Map<string, (event?: service.ICustomEvent<any>) => void>([
      ...this.eventHandlersMap,
      ['dataSelected', (event) => console.log(event)],
    ]);

    // Prepare status message for success
    this.prepareStatusMessage(successElement, successClass);
    this.displayMessage = 'Data Selected event set successfully. Select data to see event in console.';
  }

  async print() {
    // Get report from the wrapper component
    const report: Report = this.reportObj.getReport();
    try {
      await report.print();
    }
    catch (errors) {
        console.log(errors);
    }
  }

  /**
   * Prepare status message while using JS SDK APIs
   * @param img Image to show with the display message
   * @param type Type of the message
   *
   * @returns void
   */
  prepareStatusMessage(img: HTMLImageElement, type: string) {
    // Prepend Image to the display message
    this.statusRef.nativeElement.prepend(img);

    // Add class to the display message
    this.statusRef.nativeElement.classList.add(type);
  }
}
