// Copyright (c) .NET Foundation. All rights reserved.
// Licensed under the MIT License.

// ----------------------------------------------------------------------------
// Connector Namespace + a single SharePoint Online connection. Access policies
// grant the namespace polling runtime, function-app managed identity, and
// optionally the deployer user permission to call the connection.
// ----------------------------------------------------------------------------

@minLength(1)
param name string

@minLength(1)
param location string

param tags object = {}

@description('Name of the SharePoint connection to create on the namespace.')
param sharepointConnectionName string

@description('Object id of the function app system-assigned managed identity.')
param functionAppPrincipalId string = ''

@description('Optional. AAD object id of a user (typically the deployer) to also grant access to the connection so the same code can be debugged locally with `az login` credentials.')
param userPrincipalId string = ''

param tenantId string = tenant().tenantId

resource connectorNamespace 'Microsoft.Web/connectorGateways@2026-05-01-preview' = {
  name: name
  location: location
  tags: tags
  identity: {
    type: 'SystemAssigned'
  }
}

resource sharepointConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = {
  parent: connectorNamespace
  name: sharepointConnectionName
  properties: {
    connectorName: 'sharepointonline'
  }
}

resource sharepointConnectionMsiAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(functionAppPrincipalId)) {
  parent: sharepointConnection
  name: 'functionapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: functionAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource sharepointConnectionNamespaceAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = {
  parent: sharepointConnection
  name: 'namespace-trigger-runtime'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: connectorNamespace.identity.principalId
        tenantId: tenantId
      }
    }
  }
}

resource sharepointConnectionUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(userPrincipalId)) {
  parent: sharepointConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

@description('The resource ID of the Connector Namespace.')
output resourceId string = connectorNamespace.id

@description('The name of the Connector Namespace.')
output name string = connectorNamespace.name

@description('The name of the SharePoint connection.')
output sharepointConnectionName string = sharepointConnection.name

@description('The connection runtime URL for the SharePoint connection.')
output sharepointConnectionRuntimeUrl string = sharepointConnection.properties.connectionRuntimeUrl
