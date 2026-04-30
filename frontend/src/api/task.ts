import client from './client'

export interface TaskClinicalInfoSnapshot {
  [key: string]: any
}

export interface TaskItem {
  id: string
  task_type: string
  status: string
  project_type: string | null
  project_name?: string | null
  sample_id?: string | null
  sample_label?: string | null
  upload_id?: string | null
  original_filename?: string | null
  clinical_info_snapshot?: TaskClinicalInfoSnapshot
  total_files: number
  completed_files: number
  failed_files: number
  output_path?: string | null
  created_at: string | null
  started_at: string | null
  completed_at: string | null
  duration_seconds: number | null
  errors: string[]
  warnings?: string[]
  editable?: boolean
}

export interface TaskListResponse {
  items: TaskItem[]
  total: number
  page: number
  page_size: number
}

export interface TaskStats {
  total: number
  completed: number
  failed: number
  running: number
  pending: number
}

export interface TaskResultItem {
  id: number
  task_id: string
  file_index: number
  excel_filename: string
  status: string
  output_path?: string | null
  duration_seconds?: number | null
  errors: string[]
  warnings: string[]
  validation_summary?: Record<string, any>
}

export const taskApi = {
  async list(params: { status?: string; task_type?: string; page?: number; page_size?: number } = {}): Promise<TaskListResponse> {
    const { data } = await client.get('/tasks', { params })
    return data.data
  },

  async getStats(): Promise<TaskStats> {
    const { data } = await client.get('/tasks/stats')
    return data.data
  },

  async get(taskId: string): Promise<TaskItem> {
    const { data } = await client.get(`/tasks/${taskId}`)
    return data.data
  },

  async getResults(taskId: string): Promise<TaskResultItem[]> {
    const { data } = await client.get(`/tasks/${taskId}/results`)
    return data.data
  },

  async cancel(taskId: string): Promise<void> {
    await client.delete(`/tasks/${taskId}`)
  },

  async remove(taskId: string): Promise<void> {
    await client.delete(`/tasks/${taskId}/record`)
  },
}
