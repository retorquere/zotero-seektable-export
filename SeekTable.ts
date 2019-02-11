declare const Zotero: any

const aliases = Object.entries({
  bookTitle: 'publicationTitle',
  thesisType: 'type',
  university: 'publisher',
  letterType: 'type',
  manuscriptType: 'type',
  interviewMedium: 'medium',
  distributor: 'publisher',
  videoRecordingFormat: 'medium',
  genre: 'type',
  artworkMedium: 'medium',
  websiteType: 'type',
  websiteTitle: 'publicationTitle',
  institution: 'publisher',
  reportType: 'type',
  reportNumber: 'number',
  billNumber: 'number',
  codeVolume: 'volume',
  codePages: 'pages',
  dateDecided: 'date',
  reporterVolume: 'volume',
  firstPage: 'pages',
  caseName: 'title',
  docketNumber: 'number',
  documentNumber: 'number',
  patentNumber: 'number',
  issueDate: 'date',
  dateEnacted: 'date',
  publicLawNumber: 'number',
  nameOfAct: 'title',
  subject: 'title',
  mapType: 'type',
  blogTitle: 'publicationTitle',
  postType: 'type',
  forumTitle: 'publicationTitle',
  audioRecordingFormat: 'medium',
  label: 'publisher',
  presentationType: 'type',
  studio: 'publisher',
  network: 'publisher',
  episodeNumber: 'number',
  programTitle: 'publicationTitle',
  audioFileType: 'medium',
  company: 'publisher',
  proceedingsTitle: 'publicationTitle',
  encyclopediaTitle: 'publicationTitle',
  dictionaryTitle: 'publicationTitle',
})

const ignore = new Set(['attachment', 'note'])

function quote(value) {
  if (typeof value === 'number') return `${value}`
  if (!value) return ''
  if (!value.match(/[,"]/)) return value
  return `"${value.replace(/"/g, '""')}"`
}

function make_row(cells) {
  return cells.map(quote).join(',') + '\r\n'
}

function makeGenerator(type) {
  const next = Zotero[`next${type}`].bind(Zotero)
  return function*() {
    let obj
    while (obj = next()) {
      yield obj
    }
  }
}

const getCollections = makeGenerator('Collection')
const getItems = makeGenerator('Item')

function doExport() {
  const collectionName = {}
  for (const collection of getCollections()) {
    collectionName[collection.primary.key] = collection.name
  }

  const items = []
  for (const item of getItems()) {
    if (ignore.has(item.itemType)) continue

    for (const [alias, field] of aliases) {
      if (item[alias]) {
        item[field] = item[alias]
        delete item[alias]
      }
    }

    delete item.attachments
    delete item.uri
    delete item.relations
    delete item.version

    const creators = (item.creators || []).map(creator => {
      let name = ''
      if (creator.name) name = creator.name
      if (creator.lastName) name = creator.lastName
      if (creator.firstName) name += (name ? ', ' : '') + creator.firstName
      return name
    })
    if (!creators.length) creators.push('')
    delete item.creators

    item.notes = item.notes ? item.notes.map(note => `<div>${note.note}</div>`).join('\n') : ''

    const tags = (item.tags || []).map(tag => tag.tag)
    if (!tags.length) tags.push('')
    delete item.tags

    const collections = Array.from(new Set((item.collections || []).map(key => collectionName[key]).filter(coll => coll)))
    if (!collections.length) collections.push('')
    delete item.collections

    for (const creator of creators) {
      for (const collection of collections) {
        for (const tag of tags) {
          items.push({...item, creator, tag, collection})
        }
      }
    }
  }

  const headers = []
  for (const item of items) {
    for (const header of Object.keys(item)) {
      if (!headers.includes(header)) headers.push(header)
    }
  }
  headers.sort()
  Zotero.write(make_row(headers))

  for (const item of items) {
    Zotero.write(make_row(headers.map(header => item[header])))
  }
}

/*
{
   "relations" : {},
   "attachments" : [
      {
         "contentType" : "application/pdf",
         "itemType" : "attachment",
         "parentItem" : "VUL8ZVJ8",
         "uri" : "http://zotero.org/users/local/6z7M0kXV/items/UTUXSHXA",
         "filename" : "Araz et al. - 2014 - Using Google Flu Trends data in forecasting influe.pdf",
         "tags" : [],
         "version" : 0,
         "charset" : "",
         "localPath" : "/home/emile/.BBTZ5TEST/zotero/storage/UTUXSHXA/Araz et al. - 2014 - Using Google Flu Trends data in forecasting influe.pdf",
         "dateModified" : "2019-01-16T14:48:43Z",
         "linkMode" : "imported_file",
         "relations" : {},
         "title" : "Araz et al. - 2014 - Using Google Flu Trends data in forecasting influe.pdf",
         "dateAdded" : "2019-01-16T14:48:43Z"
      }
   ],
   "url" : "http://www.ajemjournal.com/article/S0735-6757(14)00421-5/abstract",
   "dateAdded" : "2019-01-16T11:47:49Z",
   "accessDate" : "2016-10-07T22:48:15Z",
   "publicationTitle" : "The American Journal of Emergency Medicine",
   "notes" : [
      {
         "dateModified" : "2019-01-16T14:40:52Z",
         "relations" : {},
         "dateAdded" : "2019-01-16T14:40:43Z",
         "itemType" : "note",
         "note" : "<p>stuf with <strong>bold</strong></p>",
         "parentItem" : "VUL8ZVJ8",
         "tags" : [],
         "key" : "DUSDL5F2",
         "version" : 0
      }
   ],
   "version" : 0,
   "ISSN" : "0735-6757, 1532-8171",
   "collections" : [],
   "extra" : "PMID: 25037278",
   "title" : "Using Google Flu Trends data in forecasting influenza-like-illness related ED visits in Omaha, Nebraska",
   "language" : "English",
   "pages" : "1016–1023",
   "dateModified" : "2019-01-16T14:03:46Z",
   "abstractNote" : "Introduction\nEmergency department (ED) visits ... and overcrowding.",
   "date" : "2014-09-01",
   "DOI" : "10.1016/j.ajem.2014.05.052",
   "journalAbbreviation" : "The American Journal of Emergency Medicine",
   "tags" : [
      {
         "tag" : "qwef"
      }
   ],
   "uri" : "http://zotero.org/users/local/6z7M0kXV/items/VUL8ZVJ8",
   "issue" : "9",
   "creators" : [
      {
         "lastName" : "Araz",
         "creatorType" : "author",
         "firstName" : "Ozgur M."
      },
      {
         "creatorType" : "author",
         "firstName" : "Dan",
         "lastName" : "Bentley"
      },
      {
         "creatorType" : "author",
         "firstName" : "Robert L.",
         "lastName" : "Muelleman"
      }
   ],
   "itemType" : "journalArticle",
   "volume" : "32",
   "libraryCatalog" : "www.ajemjournal.com"
}

{
   "childCollections" : {},
   "descendents" : [
      {
         "parent" : 4,
         "name" : "two",
         "type" : "collection",
         "key" : "JP8VQSQY",
         "children" : [
            {
               "id" : 2,
               "type" : "item",
               "key" : "3KWVWYSX",
               "parent" : 3
            },
            {
               "parent" : 3,
               "key" : "N3MW282E",
               "type" : "item",
               "id" : 3
            }
         ],
         "id" : 3,
         "level" : 1
      }
   ],
   "id" : 4,
   "name" : "een",
   "childItems" : [],
   "primary" : {
      "libraryID" : 1,
      "collectionID" : 4,
      "key" : "3UPQMN2N"
   },
   "fields" : {
      "parentKey" : false,
      "name" : "een"
   }
}

*/
