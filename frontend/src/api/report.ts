import client from './client'

export interface GenerateRequest {
  upload_id: string
  clinical_info: Record<string, any>
  project_type?: string | null
  project_name?: string | null
  template_name?: string | null
  strict_mode?: boolean
  template_contract_mode?: string
}

export interface GenerateResult {
  task_id: string
  success: boolean
  output_file: string | null
  duration_seconds: number | null
  errors: string[]
  warnings: string[]
}

export interface TaskStatus {
  id: string
  task_type: string
  status: string
  project_type: string | null
  total_files: number
  completed_files: number
  failed_files: number
  output_path: string | null
  created_at: string | null
  duration_seconds: number | null
  errors: string[]
  warnings: string[]
}

function getFilenameFromDisposition(disposition?: string): string | null {
  if (!disposition) return null

  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i)
  if (utf8Match?.[1]) {
    return decodeURIComponent(utf8Match[1].replace(/"/g, ''))
  }

  const plainMatch = disposition.match(/filename="?([^";]+)"?/i)
  return plainMatch?.[1] ? decodeURIComponent(plainMatch[1]) : null
}

function triggerBrowserDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export const reportApi = {
  async generate(req: GenerateRequest): Promise<GenerateResult> {
    const { data } = await client.post('/reports/generate', req)
    return data.data
  },

  async getTaskStatus(taskId: string): Promise<TaskStatus> {
    const { data } = await client.get(`/reports/${taskId}`)
    return data.data
  },

  async download(taskId: string): Promise<void> {
    const response = await client.get(`/reports/${taskId}/download`, {
      responseType: 'blob',
    })
    const filename =
      getFilenameFromDisposition(response.headers['content-disposition']) || `${taskId}.docx`
    const blob = new Blob([response.data as BlobPart], {
      type:
        response.headers['content-type'] ||
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })
    triggerBrowserDownload(blob, filename)
  },

  getDownloadUrl(taskId: string): string {
    return `/api/v1/reports/${taskId}/download`
  },
}
