import { GraphError } from '@microsoft/microsoft-graph-client';
import { ExternalConnectors } from '@microsoft/microsoft-graph-types';
import { config } from './config.js';
import { client } from './graphClient.js';
import fs from 'fs';

// Represents the document to import
interface Document {
  //Sample Id
  sampleId: string;
  // Document title
  title: string;
  // Document content. Can be plain-text or HTML
  shortDescription: string;
  //Author name
  authors: any[],
  //Author Picture
  authorsPictures: string,
  //Image URL
  imageUrl: string,
  // URL to the document in the external system
  url: string;
  // URL to the document icon. Required by Microsoft Copilot for Microsoft 365
  iconUrl: string;
  // Date and time when the document was last updated
  createdDateTime: string;
  //Last updated date
  lastModifiedDateTime: string;
  // Products List
  products: string[];
  //Other Metadata
  metadata: any[];
}

async function extract(fromCache: boolean, sinceDate: Date): Promise<Document[]> {
  // return the documents to import

  const samples: Document[] = [];

  if (fromCache === true) {
    console.log(`Loading from cache, including samples since ${sinceDate}...`);
    const cacheString = fs.readFileSync('cache.json', 'utf8');
    const cache = JSON.parse(cacheString);
    // samples.push(...cache.filter(sample => new Date(sample.updateDateTime) > sinceDate));
    samples.push(...cache.filter((sample: { updateDateTime: string | number | Date; }) => new Date(sample.updateDateTime) > sinceDate));
  }
  else {
    console.log(`Loading from API, including samples since ${sinceDate}...`);

    const pagination = {
      size: 50,
      index: 1
    };

    const api = 'https://m365-galleries.azurewebsites.net/Samples/searchSamples';
    const payload = {
      sort: {
        field: 'updateDateTime',
        descending: true
      },
      filter: {
        search: '',
        productId: [],
        authorId: '',
        categoryId: '',
        featuredOnly: false,
        metadata: []
      },
      pagination
    }

    let numSamplesRetrieved = 0;
    do {
      console.log(`Retrieving page ${pagination.index}...`);
      numSamplesRetrieved = 0;

      const response = await fetch(api, {
        method: 'POST',
        headers: {
          'content-type': 'application/json'
        },
        body: JSON.stringify(payload)
      });
      const data = await response.json();

      if (data.items.length > 0) {
        const samplesToInclude = data.items.filter((sample: { updateDateTime: string | number | Date }) => new Date(sample.updateDateTime) > sinceDate);
        samples.push(...samplesToInclude);
        numSamplesRetrieved = samplesToInclude.length;
      }

      console.log(`  ${numSamplesRetrieved} samples retrieved`);
      pagination.index++;
    }
    while (numSamplesRetrieved > 0);

    // cache the results
    fs.writeFileSync('cache.json', JSON.stringify(samples, null, 2));
  }

  return samples;
}

function getDocId(doc: Document): string {
  // Generate a unique ID for the document.
  // ID can't contain /
  // Generate an ID that you can resolve back to the document's URL
  // so that URL to item resolvers can properly record activity.
  return doc.sampleId;
}

function getLastCrawledSampleDate() {
  let lastCrawledSampleDate: Date = new Date(0);
  try {
    const lastCrawledSampleDateStr = fs.readFileSync('latestChange.txt', 'utf8');
    lastCrawledSampleDate = new Date(lastCrawledSampleDateStr);
    if (isNaN(lastCrawledSampleDate.getTime())) {
      lastCrawledSampleDate = new Date(0);
    }
  }
  catch (e) {
    // ignore
  }
  return lastCrawledSampleDate;
}

function transform(documents: Document[]): ExternalConnectors.ExternalItem[] {
  return documents.map(doc => {
    const docId = getDocId(doc);
    return {
      id: docId,
      properties: {
        // Add properties as defined in the schema in config.ts
        title: doc.title ?? '',
        description: doc.shortDescription ?? '',
        'authors@odata.type': 'Collection(String)',
        authors: doc.authors.map(author => author.displayName),
        'authorsPictures@odata.type': 'Collection(String)',
        authorsPictures: doc.authors.map(author => author.pictureUrl),
        imageUrl: "",
        url: `https://adoption.microsoft.com/sample-solution-gallery/sample/${docId}/`,
        iconUrl: 'https://raw.githubusercontent.com/pnp/media/master/pnp-logos-generics/png/teal/300w/pnp-samples-teal-300.png',
        createdDateTime: doc.createdDateTime,
        lastModifiedDateTime: doc.lastModifiedDateTime,
        'products@odata.type': 'Collection(String)',
        products: doc.products,
        'metadata@odata.type': 'Collection(String)',
        metadata: doc.metadata.map(m => `${m.key}=${m.value}`)
      },
      content: {
        value: doc.shortDescription ?? '',
        type: 'text'
      },
      acl: [
        {
          accessType: 'grant',
          type: 'everyone',
          value: 'everyone'
        }
      ]
    } as ExternalConnectors.ExternalItem
  });
}

async function load(externalItems: ExternalConnectors.ExternalItem[]) {
  const { id } = config.connection;
  for (const doc of externalItems) {
    try {
      console.log(`Loading ${doc.id}...`);
      await client
        .api(`/external/connections/${id}/items/${doc.id}`)
        .header('content-type', 'application/json')
        .put(doc);
      console.log('  DONE');
    }
    catch (e) {
      const graphError = e as GraphError;
      console.error(`Failed to load ${doc.id}: ${graphError.message}`);
      if (graphError.body) {
        console.error(`${JSON.parse(graphError.body)?.innerError?.message}`);
      }
      return;
    }
  }
}

export async function loadContent() {

  const lastCrawledSampleDate = getLastCrawledSampleDate();

  const content = await extract(true, lastCrawledSampleDate);
  if (content.length === 0) {
    console.log(`No new samples to load`);
    return;
  }

  const transformed = transform(content);
  await load(transformed);
}

loadContent();