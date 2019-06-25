const axios = require('axios')
const { curry, forEach } = require('lodash')
const fs = require('fs')

// Simple sleep function
function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

const headers = { Authorization: `Bearer ${process.argv[2]}` }
const fetch = (url, opts = {}) =>
  axios(url, Object.assign({}, { headers }, opts))

const logErr = (err) => {
  console.error(err)
  process.exitCode = 1
  return
}

// fetchPrivateFolder :: (number) => Promise<*>
const fetchPrivateFolder = (id) =>
  fetch(`https://platform.quip.com/1/folders/${id}`)

// fetchDocs :: (Array<number>, string) => Promise<*>
const fetchDocs = (children, folderName = 'output') => {
  fs.mkdir(folderName, 0o777, (err) => {
    if (err) return logErr(`❌ Failed to create folder ${folderName}. ${err}`)

    console.log(`🗂 ${folderName} created successfully`)
  })

  const ids = children
    .filter(({ thread_id }) => !!thread_id)
    .map(({ thread_id }) => thread_id)
    .join(',')

  const folderIds = children
    .filter(({ folder_id }) => !!folder_id)
    .map(({ folder_id }) => folder_id)

  forEach(folderIds, (folderId) => fetchThreads(folderId, folderName))

  if(ids=="" || ids== null){
    // In case of empty folders
    return
  }

  return fetch(`https://platform.quip.com/1/threads/?ids=${ids}`)
    .then(writeFiles(folderName))
}

// fetchThreads :: (number, string) => Promise<*>
const fetchThreads = (folderId, parentDir) => {
  return fetch(`https://platform.quip.com/1/folders/${folderId}`)
    .then(({ data }) => {
      forEach(data, (folder) => {
        if (!folder.title) return

        fetchDocs(data.children, `${parentDir}/${folder.title}`)
      })
    })
}

// writeFiles :: Object => void
const writeFiles = curry((folderName, { data }) => {

  forEach(data, (({ thread, html}) => {

    // Replace invalid characters with '_' on Windows. e.g. '/',':' 
    const file = thread.title.replace(/\/|:/g, '_')
    const fileName = `${folderName}/${file}`
  

    if (fs.existsSync(`${fileName}.docx`)) {
      //file exists
      console.log(`✅ ${fileName}.docx already exists!`)
    } else {
      // Massive export requests will lead to HTTP 503 error, so we wait for a second here.
      sleep(1000);

      // Option 'arraybuffer' is very important here
      fetch(`https://platform.quip.com/1/threads/${thread.id}/export/docx`, { responseType:"arraybuffer" } )
      .then(writeDocx(fileName))
    }

  }))
})


const writeDocx = curry((fileName, { data }) => {
  fs.writeFile(`${fileName}.docx`, data, 'binary', (err) => {
    if (err) {
      return logErr(`Failed to save ${fileName}.docx. ${err}`)
    }
    console.log(`✅ ${fileName}.docx saved successfully`)
  })
})
// main :: () => void
const main = () => {
  if (!process.argv[2]) {
    console.log('❌ Please provide your Quip API token. Exiting.')
    process.exitCode = 1
    return
  }

  return fetch('https://platform.quip.com/1/users/current')
    .then((res) => {
      return new Promise((resolve, reject) => {
        if (res.status !== 200) return reject(`❌ Error: ${res.statusText}`)

        resolve(res.data.private_folder_id)
      })
    })
    .then(fetchPrivateFolder)
    .then(({ data: { children } }) => children)
    .then(fetchDocs)
    .catch(logErr)
}

main()
