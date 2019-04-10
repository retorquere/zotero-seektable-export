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

function make_cell(value) {
  if (typeof value === 'number') value = `${value}`
  if (!value) value = ''

  const entity = {
    '"': '&quot;',
    "'": '&apos;',
    '<': '&lt;',
    '>': '&gt;',
    '&': '&amp;',
  }

  value = value.split('\n').map(text => `<text:p>${text.replace(/[<>"'&]/g, c => entity[c])}</text:p>`).join('')
  return `<table:table-cell office:value-type="string" calcext:value-type="string">${value}</table:table-cell>`
}

function make_row(cells) {
  return `<table:table-row>${cells.map(make_cell).join('')}</table:table-row>`
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

    item.year = null
    if (item.date) {
      const date = Zotero.Utilities.strToDate(item.date)
      if (date) item.year = date.year
      item.date = Zotero.Utilities.strToISO(item.date) || item.date
    }

    const collection_paths = (item.collections || []).map(key => collections[key] ? collections[key].path.join(', ') : '').filter(coll => coll)
    if (!collection_paths.length) collection_paths.push('')
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
      item.automatic_tags = tags.automatic.join(bundle_automatic_tags_sep)

      for (const collection of collection_paths) {
        for (const tag of tags.manual) {
          items.push({...item, tag, collection})
        }
      }

    } else {

      for (const collection of collection_paths) {
        for (const tag of tags.manual) {
          for (const automatic_tag of tags.automatic) {
            items.push({...item, tag, automatic_tag, collection})
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

  Zotero.write(`
    <?xml version='1.0' encoding='UTF-8'?>
    <office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
    <office:body>
      <office:spreadsheet>
      <table:calculation-settings table:automatic-find-labels="false" table:use-regular-expressions="false" table:use-wildcards="true"/>
      <table:table table:name="My Library">
  `.trim())
  for (const header of headers) {
    Zotero.write('<table:table-column/>')
  }

  Zotero.write(make_row(headers))

  for (const item of items) {
    Zotero.write(make_row(headers.map(header => item[header])))
  }

  Zotero.write(`
          </table:table>
        </office:spreadsheet>
      </office:body>
    </office:document>
  `)
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
