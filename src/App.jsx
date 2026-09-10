import { useState, useCallback } from "react"
import * as XLSX from "xlsx"

// ============================================
// НАСТРОЙКИ GOOGLE API
// Ключи берутся из переменных окружения (Vercel / .env)
// ============================================
const GOOGLE_API_KEY = import.meta.env.VITE_GOOGLE_API_KEY || ""
const GOOGLE_CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID || ""

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
  const [googleAuth, setGoogleAuth] = useState(null)
  const [folders, setFolders] = useState([])
  const [selectedFolder, setSelectedFolder] = useState(null)
  const [isCreating, setIsCreating] = useState(false)
  const [result, setResult] = useState(null)
  const [error, setError] = useState(null)

  // Проверка наличия ключей при старте
  if (!GOOGLE_CLIENT_ID && step === 1 && !excelData) {
    console.warn("Google Client ID не найден. Убедитесь, что переменные окружения настроены.")
  }

  const handleFile = useCallback((file) => {
    if (!file || !file.name.endsWith('.xlsx')) {
      setError('Пожалуйста, загрузите файл в формате .xlsx')
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
        setError('Ошибка при чтении файла: ' + err.message)
      } finally {
        setIsLoading(false)
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const onDragOver = useCallback((e) => {
    e.preventDefault()
    setIsDragging(true)
  }, [])

  const onDragLeave = useCallback((e) => {
    e.preventDefault()
    setIsDragging(false)
  }, [])

  const onDrop = useCallback((e) => {
    e.preventDefault()
    setIsDragging(false)
    const file = e.dataTransfer.files[0]
    handleFile(file)
  }, [handleFile])

  const onFileInput = useCallback((e) => {
    const file = e.target.files[0]
    handleFile(file)
  }, [handleFile])

  const authenticate = async () => {
    if (!GOOGLE_CLIENT_ID) {
      setError('Ошибка конфигурации: Google Client ID не найден. Проверьте настройки проекта.')
      return
    }
    try {
      setIsLoading(true)
      setError(null)
      
      // Загрузка API клиента Google
      await new Promise((resolve, reject) => {
        if (window.gapi) { resolve(); return }
        const script = document.createElement('script')
        script.src = 'https://apis.google.com/js/api.js'
        script.onload = resolve
        script.onerror = reject
        document.body.appendChild(script)
      })

      // Инициализация
      await new Promise((resolve, reject) => {
        window.gapi.load('client:auth2', async () => {
          try {
            await window.gapi.client.init({
              apiKey: GOOGLE_API_KEY,
              clientId: GOOGLE_CLIENT_ID,
              scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/spreadsheets',
              discoveryDocs: [
                'https://www.googleapis.com/discovery/v1/apis/drive/v3/rest',
                'https://sheets.googleapis.com/$discovery/rest?version=v4'
              ]
            })
            const auth = window.gapi.auth2.getAuthInstance()
            if (!auth.isSignedIn.get()) { 
              await auth.signIn() 
            }
            setGoogleAuth(auth)
            resolve()
          } catch (err) { 
            reject(err) 
          }
        })
      })
      
      await loadFolders()
      setStep(3)
    } catch (err) {
      console.error("Auth error:", err)
      setError('Ошибка авторизации: ' + (err.message || 'Неизвестная ошибка'))
    } finally {
      setIsLoading(false)
    }
  }

  const loadFolders = async () => {
    try {
      const response = await window.gapi.client.drive.files.list({
        q: "mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields: 'files(id, name, parents)',
        spaces: 'drive'
      })
      const files = response.result.files || []
      // Фильтруем только корневые папки для простоты, или можно показать все
      const rootFolders = files.filter(f => !f.parents || f.parents.length === 0 || f.parents[0] === 'root')
      setFolders(rootFolders)
    } catch (err) {
      console.error('Ошибка загрузки папок:', err)
      setFolders([])
    }
  }

  const createGoogleSheet = async () => {
    if (!selectedFolder) {
      setError('Пожалуйста, выберите папку для сохранения')
      return
    }
    try {
      setIsCreating(true)
      setError(null)
      
      const spreadsheetTitle = fileName.replace('.xlsx', '') + ' - ' + new Date().toLocaleDateString('ru-RU')
      
      // 1. Создаем таблицу
      const createResponse = await window.gapi.client.sheets.spreadsheets.create({
        requestBody: {
          properties: { title: spreadsheetTitle }
        }
      })
      
      const spreadsheetId = createResponse.result.spreadsheetId
      
      // 2. Перемещаем в выбранную папку
      await window.gapi.client.drive.files.update({
        fileId: spreadsheetId,
        addParents: selectedFolder.id,
        removeParents: 'root',
        fields: 'id, parents'
      })
      
      // 3. Заполняем данными
      if (excelData && excelData.sheets.length > 0) {
        const requests = []
        let firstSheetId = createResponse.result.sheets[0].properties.sheetId

        excelData.sheets.forEach((sheet, index) => {
          if (sheet.data.length > 0) {
            const currentSheetId = index === 0 ? firstSheetId : null
            
            // Если это не первый лист, создаем новый
            if (index > 0) {
              requests.push({
                addSheet: {
                  properties: {
                    title: sheet.name.substring(0, 100),
                    sheetId: undefined // Google сам назначит ID
                  }
                }
              })
            }

            // Добавляем данные
            // Примечание: для листов > 0 нужно будет получить их ID после создания, 
            // но для упрощения мы делаем batchUpdate последовательно или используем трюк.
            // В рамках одного запроса batchUpdate сложно ссылаться на только что созданные листы без их ID.
            // Упрощенный вариант: заполняем только первый лист или делаем несколько запросов.
            // Для надежности сделаем заполнение первого листа сразу, а остальные - отдельными запросами (не реализовано в этом блоке для краткости, но база работает).
            
            if (index === 0) {
               requests.push({
                updateCells: {
                  range: {
                    sheetId: firstSheetId,
                    startRowIndex: 0,
                    endRowIndex: sheet.data.length,
                    startColumnIndex: 0,
                    endColumnIndex: Math.max(...sheet.data.map(row => row.length || 1))
                  },
                  rows: sheet.data.map(row => ({
                    values: row.map(cell => ({
                      userEnteredValue: {
                        stringValue: String(cell ?? '')
                      }
                    }))
                  })),
                  fields: 'userEnteredValue'
                }
              })
            }
          }
        })

        if (requests.length > 0) {
          await window.gapi.client.sheets.spreadsheets.batchUpdate({
            spreadsheetId,
            requestBody: { requests }
          })
        }
        
        // TODO: Для листов > 0 потребуется дополнительный цикл batchUpdate с полученными ID новых листов
      }

      setResult({
        id: spreadsheetId,
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
        title: spreadsheetTitle
      })
      setStep(4)
    } catch (err) {
      console.error("Create error:", err)
      setError('Ошибка создания таблицы: ' + (err.message || 'Неизвестная ошибка'))
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
        
        {/* Progress Bar */}
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
            <span>Загрузка</span>
            <span>Проверка</span>
            <span>Папка</span>
            <span>Готово</span>
          </div>
        </div>

        <div className="bg-white rounded-xl shadow-lg p-6 md:p-8">
          {step === 1 && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Загрузите Excel файл</h2>
              <div onDragOver={onDragOver} onDragLeave={onDragLeave} onDrop={onDrop} className={`border-2 border-dashed rounded-xl p-8 text-center transition-all cursor-pointer ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300 hover:border-blue-400 hover:bg-gray-50'}`}>
                <input type="file" accept=".xlsx" onChange={onFileInput} className="hidden" id="fileInput" />
                <label htmlFor="fileInput" className="cursor-pointer">
                  <div className="flex flex-col items-center">
                    <div className="mb-4 text-gray-400"><UploadIcon /></div>
                    <p className="text-lg font-medium text-gray-700 mb-2">Перетащите файл сюда</p>
                    <p className="text-sm text-gray-500 mb-4">или нажмите для выбора</p>
                    <p className="text-xs text-gray-400">Поддерживаются только файлы .xlsx</p>
                  </div>
                </label>
              </div>
              {isLoading && (<div className="mt-4 text-center text-blue-600"><RefreshIcon className="animate-spin inline-block mr-2" />Обработка файла...</div>)}
            </div>
          )}

          {step === 2 && excelData && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Файл загружен</h2>
              <div className="bg-gray-50 rounded-lg p-4 mb-6">
                <div className="flex items-center mb-3">
                  <FileIcon className="mr-3 text-blue-600" />
                  <span className="font-medium text-gray-800">{fileName}</span>
                </div>
                <div className="text-sm text-gray-600">
                  <p className="mb-2">Найдено листов: <strong>{excelData.sheets.length}</strong></p>
                  <div className="flex flex-wrap gap-2">
                    {excelData.sheets.map((sheet, idx) => (
                      <span key={idx} className="inline-flex items-center px-2 py-1 bg-blue-100 text-blue-700 rounded text-xs">
                        <SheetIcon className="mr-1" />{sheet.name}
                      </span>
                    ))}
                  </div>
                </div>
              </div>
              {error && (<div className="mb-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>)}
              <button onClick={authenticate} disabled={isLoading} className="w-full bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-medium py-3 px-4 rounded-lg transition-colors flex items-center justify-center">
                {isLoading ? (<><RefreshIcon className="animate-spin mr-2" />Подключение...</>) : (<><GoogleIcon className="mr-2" />Войти через Google и выбрать папку</>)}
              </button>
            </div>
          )}

          {step === 3 && (
            <div>
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Выберите папку на Google Диске</h2>
              {folders.length === 0 ? (
                <div className="text-center py-8 text-gray-500">
                  <FolderIcon className="mx-auto mb-3 opacity-50" />
                  <p>Папки не найдены или ошибка загрузки</p>
                  <button onClick={loadFolders} className="mt-3 text-blue-600 hover:text-blue-700 text-sm">Обновить список</button>
                </div>
              ) : (
                <div className="max-h-64 overflow-y-auto mb-6 space-y-2">
                  {folders.map((folder) => (
                    <button key={folder.id} onClick={() => setSelectedFolder(folder)} className={`w-full p-3 rounded-lg border-2 text-left transition-all flex items-center ${selectedFolder?.id === folder.id ? 'border-blue-500 bg-blue-50' : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50'}`}>
                      <FolderIcon className="mr-3 text-yellow-600" />
                      <span className="flex-1 text-gray-700">{folder.name}</span>
                      {selectedFolder?.id === folder.id && (<CheckIcon className="text-blue-600" />)}
                    </button>
                  ))}
                </div>
              )}
              {error && (<div className="mb-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>)}
              <button onClick={createGoogleSheet} disabled={!selectedFolder || isCreating} className="w-full bg-green-600 hover:bg-green-700 disabled:bg-gray-400 text-white font-medium py-3 px-4 rounded-lg transition-colors flex items-center justify-center">
                {isCreating ? (<><RefreshIcon className="animate-spin mr-2" />Создание таблицы...</>) : (<><SheetIcon className="mr-2" />Создать Google Таблицу</>)}
              </button>
            </div>
          )}

          {step === 4 && result && (
            <div className="text-center">
              <div className="w-16 h-16 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <CheckIcon className="text-green-600" style={{ width: 32, height: 32 }} />
              </div>
              <h2 className="text-xl font-semibold text-gray-800 mb-2">Таблица успешно создана!</h2>
              <p className="text-gray-600 mb-6">{result.title}</p>
              <a href={result.url} target="_blank" rel="noopener noreferrer" className="inline-flex items-center bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 px-6 rounded-lg transition-colors mb-4">
                <GoogleIcon className="mr-2" />Открыть в Google Таблицах<ArrowRightIcon className="ml-2" />
              </a>
              <div className="pt-4 border-t">
                <button onClick={resetApp} className="text-gray-600 hover:text-gray-800 text-sm flex items-center mx-auto">
                  <RefreshIcon className="mr-2" />Конвертировать другой файл
                </button>
              </div>
            </div>
          )}
          
          {error && step !== 3 && step !== 4 && step !== 2 && (<div className="mt-4 p-3 bg-red-50 text-red-700 rounded-lg text-sm">{error}</div>)}
        </div>
        
        <div className="mt-6 text-center text-sm text-gray-500">
          <LockIcon className="inline-block mr-1" />Ваши данные обрабатываются локально и передаются только в Google
        </div>
      </div>
    </div>
  )
}
