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

function assign_path(collection, collections) {
  if (!collection) return []

  if (collection.parent && !collections[collection.parent]) collection.parent = false
  collection.path = collection.path || assign_path(collections[collection.parent], collections).concat(collection.name)

  return collection.path.slice()
}

function doExport() {
  const collections = {}
  for (const collection of getCollections()) {
    const children = collection.children || collection.descendents || []
    const key = (collection.primary ? collection.primary : collection).key

    collections[key] = {
      id: collection.id,
      key,
      parent: collection.fields.parentKey,
      name: collection.name,
      items: collection.childItems,
      collections: children.filter(coll => coll.type === 'collection').map(coll => coll.key),
    }
  }
  for (const collection of Object.values(collections)) {
    assign_path(collection, collections)
  }
  for (const key of Object.keys(collections)) {
    collections[key] = collections[key].path
  }

  const bundle_automatic_tags = Zotero.getOption('Bundle automatic tags')
  let bundle_automatic_tags_sep = ', '
  if (bundle_automatic_tags) {
    bundle_automatic_tags_sep = Zotero.getHiddenPref('SeekTable.delimiter.automatic_tags')
    if (!bundle_automatic_tags_sep || bundle_automatic_tags_sep.toLowerCase() === 'comma') {
      bundle_automatic_tags_sep = ', '

    } else if (bundle_automatic_tags_sep.match(/^(cr|lf)+$/i)) {
      bundle_automatic_tags_sep = bundle_automatic_tags_sep.replace(/cr/ig, '\r').replace(/lf/ig, '\n')

    } else {
        Zotero.debug(`using verbatim SeekTable.delimiter.automatic_tags ${JSON.stringify(bundle_automatic_tags_sep)}`)

    }
  }

  const items = []
  for (const item of getItems()) {
    if (item.itemType === 'attachment') continue

    if (item.itemType === 'note') {
      item.notes = item.note ? [ { note: item.note } ] : []
      delete item.notes
    }

    for (const [alias, field] of aliases) {
      if (typeof item[alias] !== 'undefined') {
        item[field] = item[alias]
        delete item[alias]
      }
    }

    delete item.attachments
    delete item.uri
    delete item.relations
    delete item.version
    delete item.citekey
    delete item.itemID
    delete item.key
    delete item.libraryID

    item.creators = (item.creators || []).map(creator => creator.name ? creator.name : [creator.lastName, creator.firstName].filter(name => name).join(', ')).join('; ')

    if (!item.url && item.DOI) item.url = `${item.DOI.startsWith('http') ? '' : 'http://doi.org/'}${item.DOI}`

    if (!item.notes) {
      item.notes = ''
    } else if (item.notes.length === 1) {
      item.notes = item.notes[0].note
    } else {
      item.notes = item.notes.map(note => `<div>${note.note}</div>`).join('')
    }
    item.notes = item.notes.replace(/[\r\n]+/g, ' ').replace(/\u00A0/g, ' ')

    item.extra = (item.extra || '').replace(/[\r\n]+/g, ' ').replace(/\u00A0/g, ' ')
    if (!item.url) {
      item.extra = item.extra.replace(/(?:^|\n)PMCID:\s*([^\n]+)/, (_, pmcid) => {
        item.url = `https://www.ncbi.nlm.nih.gov/pmc/articles/${pmcid}/`
        return ''
      }).strip()
    }
    if (!item.url) {
      item.extra = item.extra.replace(/(?:^|\n)PMID:\s*([^\n]+)/, (_, pmid) => {
        item.url = `https://www.ncbi.nlm.nih.gov/pubmed/${pmid}`
        return ''
      }).strip()
    }
    if (item.extra.includes(':')) item.extra = ''

    item.year = null
    if (item.date) {
      const date = Zotero.Utilities.strToDate(item.date)
      if (date) item.year = date.year
      item.date = Zotero.Utilities.strToISO(item.date) || item.date
    }

    const collection_paths = (item.collections || []).map(key => collections[key]).filter(path => path && path.length)
    if (!collection_paths.length) collection_paths.push([])
    delete item.collections

    const tags = {
      manual: (item.tags || []).filter(tag => tag.type !== 1).map(tag => tag.tag),
      automatic: (item.tags || []).filter(tag => tag.type === 1).map(tag => tag.tag),
    }
    for (const type of Object.keys(tags)) {
      if (!tags[type].length) tags[type].push('')
    }
    delete item.tags

    if (bundle_automatic_tags) {
      item.automatic_tag = tags.automatic.join(bundle_automatic_tags_sep)

      for (const collection of collection_paths) {
        for (const tag of tags.manual) {
          items.push({ ...item, tag, parent_collections: collection.join(', '), collection: collection[collection.length - 1] || '' })
        }
      }

    } else {
      item.all_automatic_tags = tags.automatic.join(bundle_automatic_tags_sep)

      for (const collection of collection_paths) {
        for (const tag of tags.manual) {
          for (const automatic_tag of tags.automatic) {
            items.push({ ...item, tag, automatic_tag, parent_collections: collection.join(', '), collection: collection[collection.length - 1] || '' })
          }
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
