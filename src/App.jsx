import { useState, useEffect } from "react"
import * as XLSX from "xlsx"

const GOOGLE_API_KEY = import.meta.env.VITE_GOOGLE_API_KEY || ""
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || ""

// Иконки
const FileIcon = () => <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
const UploadIcon = () => <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
const GoogleIcon = () => <svg width="20" height="20" viewBox="0 0 24 24"><path fill="#4285F4" d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"/><path fill="#34A853" d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"/><path fill="#FBBC05" d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"/><path fill="#EA4335" d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"/></svg>
const FolderIcon = () => <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>
const CheckIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>
const LockIcon = () => <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
const SheetIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>
const ArrowRightIcon = () => <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/></svg>
const RefreshIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>

export default function App() {
  const [step, setStep] = useState(1)
  const [excelData, setExcelData] = useState(null)
  const [fileName, setFileName] = useState("")
  const [isDragging, setIsDragging] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [accessToken, setAccessToken] = useState(null)
  const [folders, setFolders] = useState([])
  const [selectedFolder, setSelectedFolder] = useState(null)
  const [isCreating, setIsCreating] = useState(false)
  const [result, setResult] = useState(null)
  const [error, setError] = useState(null)
  const [gapiReady, setGapiReady] = useState(false)
  const [gisReady, setGisReady] = useState(false)

  // Загрузка скриптов Google
  useEffect(() => {
    const loadScripts = async () => {
      try {
        // 1. Загрузка GAPI
        if (!window.gapi) {
          await new Promise((resolve, reject) => {
            const script = document.createElement('script')
            script.src = 'https://apis.google.com/js/api.js'
            script.onload = resolve
            script.onerror = () => reject(new Error("Failed to load gapi"))
            document.body.appendChild(script)
          })
        }

        // 2. Инициализация GAPI клиента
        await new Promise((resolve) => {
          window.gapi.load('client', () => {
            window.gapi.client.init({
              apiKey: GOOGLE_API_KEY,
              clientId: GOOGLE_CLIENT_ID,
              discoveryDocs: [
                'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest',
                'https://sheets.googleapis.com/$discovery/rest?version=v4'
              ]
            }).then(resolve).catch(resolve)
          })
        })
        setGapiReady(true)

        // 3. Загрузка GIS
        if (!window.google || !window.google.accounts) {
          await new Promise((resolve, reject) => {
            const script = document.createElement('script')
            script.src = 'https://accounts.google.com/gsi/client'
            script.onload = resolve
            script.onerror = () => reject(new Error("Failed to load gis"))
            document.body.appendChild(script)
          })
        }
        setGisReady(true)
      } catch (err) {
        console.error("Script load error:", err)
        setError("Ошибка загрузки сервисов Google.")
      }
    }

    loadScripts()
  }, [])

  const handleFile = (file) => {
    if (!file || !file.name.endsWith('.xlsx')) {
      setError('Пожалуйста, загрузите файл .xlsx')
      return
    }
    setError(null)
    setIsLoading(true)
    setFileName(file.name)
    
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheets = workbook.SheetNames.map(name => ({
          name,
          data: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 })
        }))
        setExcelData({ workbook, sheets })
        setStep(2)
      } catch (err) {
        setError('Ошибка чтения: ' + err.message)
      } finally {
        setIsLoading(false)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const onDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    handleFile(e.dataTransfer.files[0])
  }

  const authenticate = () => {
    if (!gisReady) {
      setError("Сервисы Google еще не загружены.")
      return
    }

    const tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets',
      callback: (response) => {
        if (response.error) {
          setError(`Auth Error: ${response.error}`)
          return
        }
        setAccessToken(response.access_token)
        if (window.gapi && window.gapi.client) {
          window.gapi.client.setToken({ access_token: response.access_token })
        }
        
        loadFolders(response.access_token)
        setStep(3)
      },
    })

    tokenClient.requestAccessToken()
  }

  const loadFolders = async (token) => {
    const authToken = token || accessToken;
    if (!authToken) {
      setError("Нет токена доступа для загрузки папок")
      return
    }
    try {
      const res = await fetch(
        `https://www.googleapis.com/drive/v3/files?q=mimeType='application/vnd.google-apps.folder'+and+trashed=false&fields=files(id,name,parents)&spaces=drive`,
        {
          headers: { 
            'Authorization': `Bearer ${authToken}`,
            'Content-Type': 'application/json'
          }
        }
      )
      
      if (!res.ok) {
        const errData = await res.json()
        throw new Error(errData.error?.message || `HTTP ${res.status}`)
      }

      const data = await res.json()
      const files = data.files || []
      const rootFolders = files.filter(f => !f.parents || f.parents.length === 0 || f.parents[0] === 'root')
      setFolders(rootFolders)
    } catch (err) {
      console.error("Folder Load Error:", err)
      setError("Не удалось загрузить папки: " + err.message)
      setFolders([])
    }
  }

  const createGoogleSheet = async () => {
    if (!accessToken) {
      setError("Нет токена доступа")
      return
    }
    
    setIsCreating(true)
    setError(null)
    
    const title = fileName.replace('.xlsx', '') + ' - ' + new Date().toLocaleDateString('ru-RU')
    
    try {
      // 1. Создаем таблицу
      const createRes = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ properties: { title } })
      })
      
      const createData = await createRes.json()
      if (createData.error) throw new Error(createData.error.message)
      
      const spreadsheetId = createData.spreadsheetId
      const firstSheetId = createData.sheets[0].properties.sheetId

      // 2. Перемещаем в папку (если выбрана)
      if (selectedFolder) {
        const moveRes = await fetch(
          `https://www.googleapis.com/drive/v3/files/${spreadsheetId}?addParents=${selectedFolder.id}&fields=id,parents`,
          {
            method: 'PATCH',
            headers: { 
              'Authorization': `Bearer ${accessToken}`,
              'Content-Type': 'application/json'
            }
          }
        )
        if (!moveRes.ok) {
          const moveErr = await moveRes.json()
          console.warn("Move warning:", moveErr)
        }
      }

      // 3. Заполняем данными
      if (excelData && excelData.sheets.length > 0) {
        const sheet = excelData.sheets[0]
        const rowData = sheet.data
        
        if (rowData && rowData.length > 0) {
          const maxCols = Math.max(...rowData.map(r => (r ? r.length : 0)), 1)
          
          const requests = [{
            updateCells: {
              range: {
                sheetId: firstSheetId,
                startRowIndex: 0,
                endRowIndex: rowData.length,
                startColumnIndex: 0,
                endColumnIndex: maxCols
              },
              rows: rowData.map(row => {
                if (!row) return { values: [] }
                return {
                  values: row.map(cell => {
                    if (cell === null || cell === undefined || cell === '') {
                      return {}
                    }
                    if (typeof cell === 'number') {
                      return { userEnteredValue: { numberValue: cell } }
                    }
                    if (typeof cell === 'boolean') {
                      return { userEnteredValue: { boolValue: cell } }
                    }
                    return { userEnteredValue: { stringValue: String(cell) } }
                  })
                }
              }),
              fields: 'userEnteredValue'
            }
          }]

          const updateRes = await fetch(
            `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`,
            {
              method: 'POST',
              headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
              },
              body: JSON.stringify({ requests })
            }
          )
          
          const updateData = await updateRes.json()
          if (updateData.error) {
            console.warn("Warning during data update:", updateData.error)
          }
        }
      }

      setResult({
        id: spreadsheetId,
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        title
      })
      setStep(4)
    } catch (err) {
      console.error("Create Error:", err)
      setError("Ошибка создания: " + err.message)
    } finally {
      setIsCreating(false)
    }
  }

  const resetApp = () => {
    setStep(1)
    setExcelData(null)
    setFileName("")
    setFolders([])
    setSelectedFolder(null)
    setResult(null)
    setError(null)
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 py-8 px-4">
      <div className="max-w-2xl mx-auto">
        <div className="text-center mb-8">
          <h1 className="text-3xl font-bold text-gray-800 mb-2">Excel → Google Sheets</h1>
          <p className="text-gray-600">Конвертируйте Excel файлы в Google Таблицы</p>
        </div>
        
        {/* Progress */}
        <div className="mb-8">
          <div className="flex justify-between mb-2">
            {[1, 2, 3, 4].map((s) => (
              <div key={s} className={`w-10 h-10 rounded-full flex items-center justify-center font-semibold transition-all ${step >= s ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-500'}`}>
                {step > s ? <CheckIcon /> : s}
              </div>
            ))}
          </div>
          <div className="h-1 bg-gray-200 rounded-full">
            <div className="h-full bg-blue-600 rounded-full transition-all duration-300" style={{ width: `${((step - 1) / 3) * 100}%` }} />
          </div>
          <div className="flex justify-between mt-2 text-xs text-gray-500">
            <span>Файл</span><span>Проверка</span><span>Папка</span><span>Готово</span>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-lg p-6 md:p-8">
          {!gapiReady && step === 1 && (
             <div className="mb-6 p-4 bg-yellow-50 border-l-4 border-yellow-400 text-yellow-700 text-sm flex items-center">
               <RefreshIcon className="animate-spin mr-2" /> Загрузка сервисов API...
             </div>
          )}

          {step === 1 && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Загрузите Excel файл</h2>
              <div 
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true) }} 
                onDragLeave={() => setIsDragging(false)} 
                onDrop={onDrop} 
                className={`border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'}`}
              >
                <input type="file" accept=".xlsx" onChange={(e) => handleFile(e.target.files[0])} className="hidden" id="fileInput" />
                <label htmlFor="fileInput" className="cursor-pointer block">
                  <div className="flex flex-col items-center">
                    <div className="mb-4 text-gray-400"><UploadIcon /></div>
                    <p className="text-lg font-medium text-gray-700 mb-2">Перетащите файл сюда</p>
                    <p className="text-sm text-gray-500 mb-4">или нажмите для выбора</p>
                    <p className="text-xs text-gray-400">Только .xlsx</p>
                  </div>
                </label>
              </div>
              {error && <div className="mt-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>}
            </div>
          )}

          {step === 2 && excelData && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Файл готов</h2>
              <div className="bg-gray-50 rounded-lg p-4 mb-6">
                <div className="flex items-center mb-3">
                  <FileIcon className="mr-3 text-blue-600" />
                  <span className="font-medium text-gray-800 truncate">{fileName}</span>
                </div>
                <div className="text-sm text-gray-600">
                  <p className="mb-2">Листов: <strong>{excelData.sheets.length}</strong></p>
                  <div className="flex flex-wrap gap-2">
                    {excelData.sheets.map((s, i) => (
                      <span key={i} className="px-2 py-1 bg-blue-100 text-blue-700 rounded text-xs flex items-center">
                        <SheetIcon className="mr-1 w-3 h-3"/> {s.name}
                      </span>
                    ))}
                  </div>
                </div>
              </div>
              {error && <div className="mb-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>}
              <button onClick={authenticate} disabled={!gapiReady || !gisReady} className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-medium py-3 px-4 rounded-lg transition-colors flex items-center justify-center shadow-md">
                <GoogleIcon className="mr-2" /> Войти через Google
              </button>
            </div>
          )}

          {step === 3 && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Куда сохранить?</h2>
              
              <button 
                onClick={() => setSelectedFolder(null)} 
                className={`w-full p-4 rounded-lg border-2 text-left transition-all mb-4 flex items-center ${selectedFolder === null ? 'border-blue-500 bg-blue-50' : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50'}`}
              >
                <div className="bg-blue-100 p-2 rounded mr-3 text-blue-600"><FolderIcon /></div>
                <div className="flex-1">
                  <div className="font-medium text-gray-800">Мой диск (Корень)</div>
                  <div className="text-xs text-gray-500">Сохранить в главную папку Диска</div>
                </div>
                {selectedFolder === null && <CheckIcon className="text-blue-600" />}
              </button>

              <div className="text-sm text-gray-500 mb-2 font-medium">Или выберите папку:</div>
              
              {folders.length === 0 ? (
                <div className="text-center py-6 text-gray-400 bg-gray-50 rounded-lg border border-dashed border-gray-300">
                  <p>Папки не найдены</p>
                  <button onClick={() => loadFolders(accessToken)} className="mt-2 text-blue-600 text-xs hover:underline">Обновить список</button>
                </div>
              ) : (
                <div className="max-h-60 overflow-y-auto mb-6 space-y-2 pr-1">
                  {folders.map((f) => (
                    <button key={f.id} onClick={() => setSelectedFolder(f)} className={`w-full p-3 rounded-lg border-2 text-left transition-all flex items-center ${selectedFolder?.id === f.id ? 'border-blue-500 bg-blue-50' : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50'}`}>
                      <FolderIcon className="mr-3 text-yellow-600" />
                      <span className="flex-1 text-gray-700 truncate">{f.name}</span>
                      {selectedFolder?.id === f.id && <CheckIcon className="text-blue-600" />}
                    </button>
                  ))}
                </div>
              )}
              
              {error && <div className="mb-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>}
              <button onClick={createGoogleSheet} disabled={isCreating} className="w-full bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white font-medium py-3 px-4 rounded-lg transition-colors flex items-center justify-center shadow-md">
                {isCreating ? <><RefreshIcon className="animate-spin mr-2" /> Создание...</> : <><SheetIcon className="mr-2" /> Создать таблицу</>}
              </button>
            </div>
          )}

          {step === 4 && result && (
            <div className="text-center">
              <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <CheckIcon className="text-green-600 w-8 h-8" />
              </div>
              <h2 className="text-xl font-semibold text-gray-800 mb-2">Готово!</h2>
              <p className="text-gray-600 mb-6">{result.title}</p>
              <a href={result.url} target="_blank" rel="noreferrer" className="inline-flex items-center bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 px-6 rounded-lg transition-colors mb-4 shadow-md">
                <GoogleIcon className="mr-2" /> Открыть таблицу <ArrowRightIcon className="ml-2 w-4 h-4"/>
              </a>
              <div className="pt-4 border-t">
                <button onClick={resetApp} className="text-gray-600 hover:text-gray-800 text-sm font-medium flex items-center mx-auto">
                  <RefreshIcon className="mr-2" /> Конвертировать другой файл
                </button>
              </div>
            </div>
          )}
        </div>
        
        <div className="mt-6 text-center text-xs text-gray-500 flex items-center justify-center">
          <LockIcon className="inline-block mr-1 w-3 h-3" /> Данные обрабатываются локально и передаются только в Google
        </div>
      </div>
    </div>
  )
}
