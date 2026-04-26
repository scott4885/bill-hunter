// Azure infrastructure for Bill Hunter + Tax Archivist.
//
// Provisions: Storage, App Insights, Linux Consumption Plan (Y1), Function App.
// Sets every env var the codebase reads so the function works on first boot.

@minLength(2)
@maxLength(40)
@description('Function App name. Must be globally unique across Azure.')
param functionAppName string

@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Microsoft Entra tenant ID.')
param azureTenantId string

@description('App registration client ID.')
param azureClientId string

@secure()
@description('App registration client secret.')
param azureClientSecret string

@secure()
@description('Anthropic API key for Claude calls.')
param anthropicApiKey string

@secure()
@description('Random hex (32 bytes) used to authenticate Graph webhook POSTs. Generate with: openssl rand -hex 32')
param webhookClientStateSecret string

@secure()
@description('Initial Graph refresh token from `npm run auth:setup`.')
param graphRefreshToken string

@secure()
@description('Discord webhook URL for the daily summary. Empty string disables the post.')
param discordWebhookUrl string = ''

@description('OneDrive root folder where xlsx files and inboxes live. No trailing slash.')
param oneDriveRoot string = '/Documents/BillHunter'

var storageAccountName = 'st${take(replace(uniqueString(resourceGroup().id, functionAppName), '-', ''), 18)}'
var planName = 'plan-${functionAppName}'
var insightsName = 'ai-${functionAppName}'

resource storage 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: storageAccountName
  location: location
  sku: { name: 'Standard_LRS' }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    supportsHttpsTrafficOnly: true
  }
}

resource insights 'Microsoft.Insights/components@2020-02-02' = {
  name: insightsName
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    Request_Source: 'rest'
  }
}

resource plan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: planName
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  kind: 'functionapp,linux'
  properties: {
    reserved: true
  }
}

resource func 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    serverFarmId: plan.id
    reserved: true
    httpsOnly: true
    siteConfig: {
      linuxFxVersion: 'NODE|20'
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      appSettings: [
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storage.name};AccountKey=${storage.listKeys().keys[0].value};EndpointSuffix=core.windows.net'
        }
        { name: 'FUNCTIONS_EXTENSION_VERSION', value: '~4' }
        { name: 'FUNCTIONS_WORKER_RUNTIME', value: 'node' }
        { name: 'WEBSITE_NODE_DEFAULT_VERSION', value: '~20' }
        { name: 'AzureWebJobsFeatureFlags', value: 'EnableWorkerIndexing' }
        { name: 'FUNCTIONS_NODE_BLOCK_ON_ENTRY_POINT_ERROR', value: 'true' }
        { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: insights.properties.ConnectionString }

        // App identity / secrets
        { name: 'AZURE_TENANT_ID', value: azureTenantId }
        { name: 'AZURE_CLIENT_ID', value: azureClientId }
        { name: 'AZURE_CLIENT_SECRET', value: azureClientSecret }
        { name: 'ANTHROPIC_API_KEY', value: anthropicApiKey }
        { name: 'WEBHOOK_CLIENT_STATE_SECRET', value: webhookClientStateSecret }
        { name: 'GRAPH_REFRESH_TOKEN', value: graphRefreshToken }
        { name: 'DISCORD_WEBHOOK_URL', value: discordWebhookUrl }

        // URL the Graph subscription posts to
        { name: 'WEBHOOK_BASE_URL', value: 'https://${functionAppName}.azurewebsites.net' }

        // OneDrive paths consumed by src/graph/excel.ts and the handlers
        { name: 'ONEDRIVE_BILLS_PATH', value: '${oneDriveRoot}/bills.xlsx' }
        { name: 'ONEDRIVE_TAX_PATH_PREFIX', value: '${oneDriveRoot}/taxes_' }
        { name: 'ONEDRIVE_TRANSACTIONS_PATH_PREFIX', value: '${oneDriveRoot}/transactions_' }
        { name: 'ONEDRIVE_RENTALS_PATH_PREFIX', value: '${oneDriveRoot}/rentals_' }
        { name: 'ONEDRIVE_ATTACHMENTS_DIR', value: '${oneDriveRoot}/attachments' }
        { name: 'ONEDRIVE_STATEMENTS_INBOX', value: '${oneDriveRoot}/statements/inbox' }
        { name: 'ONEDRIVE_STATEMENTS_PROCESSED', value: '${oneDriveRoot}/statements/processed' }
        { name: 'ONEDRIVE_RENTALS_INBOX', value: '${oneDriveRoot}/rentals/inbox' }
        { name: 'ONEDRIVE_RENTALS_PROCESSED', value: '${oneDriveRoot}/rentals/processed' }
        { name: 'ONEDRIVE_PL_DIR', value: '${oneDriveRoot}/pl' }
        { name: 'ONEDRIVE_CPA_PACKETS_DIR', value: oneDriveRoot }
      ]
    }
  }
}

output functionAppName string = func.name
output webhookUrl string = 'https://${functionAppName}.azurewebsites.net/api/webhook'
output principalId string = func.identity.principalId
output storageAccountName string = storage.name
output insightsName string = insights.name
