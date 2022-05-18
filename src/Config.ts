// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

/* eslint-disable @typescript-eslint/no-inferrable-types */

// Scope of AAD app. Use the below configuration to use all the permissions provided in the AAD app through Azure portal.
// Refer https://aka.ms/PowerBIPermissions for complete list of Power BI scopes
export const scopes: string[] = ["https://analysis.windows.net/powerbi/api/Report.Read.All", "https://analysis.windows.net/powerbi/api/Dataset.Read.All"];

// Client Id (Application Id) of the AAD app.
export const clientId: string = "065ba81b-7348-4314-8f78-16d825712581";

// Id of the workspace where the report is hosted
export const workspaceId: string = "2db3b5f6-42f4-41a3-bc22-318c368d69d7";

// Id of the report to be embedded
export const reportId: string = "1802e44c-8108-4019-936d-bcc3f54eaf72";