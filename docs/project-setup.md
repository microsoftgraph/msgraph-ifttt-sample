# Project Setup

There is a fair amount of configuration needed to setup this experience. Please follow the guidance below to get started.

## Create .env file

The `.env` file contains all the secret bits needed for the service to talk to IFTTT and the Microsoft Graph.

Create a new file in the root of the project named `.env` (the '.' prefix is required). Throughout the rest of this setup guide, you will be replacing "your-xxxx-here" items with data from following steps. Keep the document open to edit as you go.

Paste in the following:

```
# Environment Config

# store your secrets and config variables in here
# only invited collaborators will be able to see your .env values

# reference these in your code with process.env.SECRET
# note: .env is a shell file so there can't be spaces around =

PORT=8080
IFTTT_KEY="your-ifttt-key-here"

TENANT_ID="your-tenant-id-here"
CLIENT_ID="your-client-id-here"
CLIENT_SECRET="your-client-secret-here"

TEST_USER="your-test-email-here"
TEST_PWD="your-test-password-here"
```

## Configure IFTTT

The next step is to create and configure the service on the IFTTT Platform.

1. To get started on the IFTTT Platform, navigate to  https://platform.ifttt.com
1. Create an account, sign in, and click `Add service`
1. Provide a Service name and ID, then click `Create`
1. Now that your service has been created, navigate to the `API` tab.

There are two important parts in the `General` section:

    **IFTTT API URL**

    This is the hosted endpoint for your IFTTT service. For this sample we will be hosting locally and exposing the local endpoint via [ngrok](https://ngrok.com/) 
    We'll come back to this later.

    **Service key**

    The service key is needed to send requests and participate on the IFTTT platform. 

    Annotate the Service key in the `.env` file.

## App registration

This sample requires an Azure Portal App Registration to configure the Graph authentication. Head over to the Azure portal to [create a new App registration](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade).

### Client + Tenant Ids

1. From the Azure Portal App registration page, click `New registration`.
1. Input a display name for your service and select the last radio button: `Accounts in any organizational directory and personal Microsoft accounts`
1. Skip the Redirect URI for now and click `Register`.
1. You should now see the Overview page for your new app registration.

Take note of the Application (client) ID and Directory (tenant) ID and annotate them in the `.env` file.

### Authentication

From your Azure Portal App registration:

1. Navigate to the `Authentication` tab.
1. Create the following Redirect URIs:
    - `http://localhost:8080`
    - `https://ifttt.com/channels/<IFTTT service ID>/authorize`
    - (optional) `https://<ngrok forwarding url>.ngrok.io`

    *Note:* `<IFTTT service ID>` can be found on the IFTTT Platform -> Service Tab -> General section.

    *Note:* `<ngrok forwarding url>` can be found in the ngrok console output in the last step. This last Redirect URI is optional and is only useful for interacting with the test page from the ngrok forwarding address. Otherwise, localhost is sufficient for any local testing purposes.
1. Under `Implicit grant`, check both boxes (Access and ID tokens).

### Client secret

From your Azure Portal App registration:

1. Navigate to the `Certificates & secrets` tab.
1. Click `New client secret` to generate a new secret value.

Annotate the new client secret value in the `.env` file.

### Expose an API

To enable IFTTT Platform to request Microsoft Graph data on behalf of our user, we must expose a scope for this activity.

From your Azure Portal App registration:

1. Navigate to the `Expose an API` tab.
1. Click `Add a scope`, then `Save and continue`.
1. Fill in the various fields:
    - Scope name: `ifttt`
    - Who can consent? `Admins and users`
    - Admin consent display name: `IFTTT Platform`
    - Admin consent description: `Enable interaction with the IFTTT Platform`
    - User consent display name `IFTTT Platform`
    - User consent description: `Enable interaction with the IFTTT Platform`
1. Click `Add scope`.

## Test user credentials

IFTTT requires a test user to impersonate during endpoint tests. In the `.env` file, provide credentials for a valid user in your tenant. Since we configured our service to support all user types, the test user can be any public MSA.

## Setup Requirements
- Ngrok: https://ngrok.com/
- Node.js: https://nodejs.org/

> A free ngrok account will be sufficient, however, beware that ngrok will return a randomized url each time, which must be updated in the IFTTT Platform config (API/General).  If this becomes an issue, it may make development easier with a static url which is available with any paid ngrok subscription.

To setup the project, run the following commands:

1. `npm i`
1. `npm run-script build`
1. `npm start`
1. `ngrok http 8080`
1. Copy ngrok url to IFTTT API URL config here: `https://platform.ifttt.com/services/<ifttt_service_name>/api`

## Test the service

With the project setup and running, you can test the service itself by navigating to http://localhost:8080. From the test page you can sign in and manually invoke any of the Actions/Triggers.

To test integration with the IFTTT platform, navigate to each of the IFTTT Platform test pages and click `Begin test`:

- [IFTTT Endpoint tests](https://platform.ifttt.com/services/msteams/api/endpoint_tests)
- [IFTTT Authentication test](https://platform.ifttt.com/services/msteams/api/authentication_test)