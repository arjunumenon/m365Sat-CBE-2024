import { ExternalConnectors } from '@microsoft/microsoft-graph-types';

export const config = {
  connection: {
    id: 'm365cbe24grcov1',
    name: 'M365CBE2024-v1',
    description: 'This is to replicate the PnP Solution Gallery in a simple way',
    activitySettings: {
      // URL to item resolves track activity such as sharing external items.
      // The recorded activity is used to improve search relevance.
      urlToItemResolvers: [
        {
          urlMatchInfo: {
            baseUrls: [
              'https://adoption.microsoft.com'
            ],
            urlPattern: '/sample-solution-gallery/sample/(?<sampleId>[^/]+)'
          },
          itemId: '{sampleId}',
          priority: 1
        } as ExternalConnectors.ItemIdResolver
      ]
    },
    searchSettings: {
      searchResultTemplates: [
        {
          id: 'm365cbe24grcov1',
          priority: 1,
          layout: {}
        }
      ]
    },
    // https://learn.microsoft.com/graph/connecting-external-content-manage-schema
    schema: {
      baseType: 'microsoft.graph.externalItem',
      // Add properties as needed
      properties: [
        {
          name: 'title',
          type: 'string',
          isQueryable: true,
          isSearchable: true,
          isRetrievable: true,
          labels: [
            'title'
          ]
        },
        {
          name: 'description',
          type: 'String',
          isQueryable: 'true',
          isSearchable: 'true',
          isRetrievable: 'true',
        },
        {
          name: 'authors',
          type: 'StringCollection',
          isQueryable: 'true',
          isSearchable: 'true',
          isRetrievable: 'true',
          labels: [
            'authors'
          ]
        },
        {
          name: 'authorsPictures',
          type: 'StringCollection',
          isRetrievable: 'true'
        },
        {
          name: 'url',
          type: 'string',
          isRetrievable: true,
          labels: [
            'url'
          ]
        },
        {
          name: 'iconUrl',
          type: 'string',
          isRetrievable: true,
          labels: [
            'iconUrl'
          ]
        },
        {
          name: 'createdDateTime',
          type: 'DateTime',
          isQueryable: 'true',
          isRetrievable: 'true',
          isRefinable: 'true',
          labels: [
            'createdDateTime'
          ]
        },
        {
          name: 'lastModifiedDateTime',
          type: 'DateTime',
          isQueryable: 'true',
          isRetrievable: 'true',
          isRefinable: 'true',
          labels: [
            'lastModifiedDateTime'
          ]
        },
        {
          name: 'products',
          type: 'StringCollection',
          isQueryable: 'true',
          isRetrievable: 'true',
          isRefinable: 'true'
        },
        {
          name: 'metadata',
          type: 'StringCollection',
          isQueryable: 'true',
          isRetrievable: 'true',
          isRefinable: 'true',
          isExactMatchRequired: 'true'
        }
      ]
    }
  } as ExternalConnectors.ExternalConnection
};