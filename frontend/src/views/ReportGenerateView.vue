<template>
  <div>
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px">
      <h2 style="margin: 0">{{ pageTitle }}</h2>
      <el-button v-if="isReadOnly" @click="switchToEdit">切换到编辑</el-button>
    </div>

    <el-card v-if="currentTask" shadow="never" style="margin-bottom: 20px">
      <el-descriptions :column="3" border size="small">
        <el-descriptions-item label="样本编号">{{ currentTask.sample_id || '-' }}</el-descriptions-item>
        <el-descriptions-item label="任务状态">{{ taskStatusLabel }}</el-descriptions-item>
        <el-descriptions-item label="原始文件">{{ currentTask.original_filename || '-' }}</el-descriptions-item>
        <el-descriptions-item label="项目类型">{{ currentTask.project_name || currentTask.project_type || '-' }}</el-descriptions-item>
        <el-descriptions-item label="任务ID">{{ currentTask.id }}</el-descriptions-item>
        <el-descriptions-item label="生成时间">
          {{ currentTask.created_at ? new Date(currentTask.created_at).toLocaleString('zh-CN') : '-' }}
        </el-descriptions-item>
      </el-descriptions>

      <div style="margin-top: 12px; display: flex; gap: 8px; flex-wrap: wrap">
        <el-button
          v-if="currentTask.status === 'completed'"
          type="primary"
          plain
          @click="downloadReport(currentTask.id)"
        >
          下载当前报告
        </el-button>
        <el-alert
          v-if="currentTask.errors?.length"
          :title="currentTask.errors[0]"
          type="error"
          show-icon
          :closable="false"
          style="flex: 1; min-width: 320px"
        />
      </div>
    </el-card>

    <!-- Step 1: Upload Excel -->
    <el-card shadow="hover" style="margin-bottom: 20px" v-loading="taskLoading">
      <template #header><strong>1. 上传 Excel 文件</strong></template>
      <el-upload
        drag
        accept=".xlsx"
        :auto-upload="false"
        :show-file-list="false"
        :disabled="isReadOnly"
        @change="handleFileChange"
      >
        <el-icon class="el-icon--upload" :size="40"><UploadFilled /></el-icon>
        <div class="el-upload__text">拖拽文件到此处，或<em>点击上传</em></div>
        <template #tip>
          <div class="el-upload__tip">仅支持 .xlsx 格式的基因检测 Excel 文件</div>
        </template>
      </el-upload>

      <div v-if="excelStore.upload" style="margin-top: 16px">
        <el-descriptions :column="3" border size="small">
          <el-descriptions-item label="文件名">{{ excelStore.upload.original_filename }}</el-descriptions-item>
          <el-descriptions-item label="大小">{{ (excelStore.upload.file_size_bytes / 1024).toFixed(1) }} KB</el-descriptions-item>
          <el-descriptions-item label="Sheet 数量">{{ excelStore.upload.sheet_names.length }}</el-descriptions-item>
          <el-descriptions-item label="检测项目类型">
            <el-tag v-if="excelStore.upload.detected_project_type" type="success">
              {{ excelStore.upload.detected_project_name || excelStore.upload.detected_project_type }}
            </el-tag>
            <el-tag v-else type="warning">未识别</el-tag>
          </el-descriptions-item>
        </el-descriptions>

        <!-- Validation warnings -->
        <div v-if="excelStore.upload.validation_warnings?.length" style="margin-top: 12px">
          <el-alert
            v-for="(w, i) in excelStore.upload.validation_warnings"
            :key="i"
            :title="w.message"
            :type="w.level === 'error' ? 'error' : w.level === 'warning' ? 'warning' : 'info'"
            show-icon
            :closable="false"
            style="margin-bottom: 4px"
          />
        </div>

        <!-- Sheet tabs preview -->
        <el-tabs v-if="excelStore.sheets.length > 0" style="margin-top: 12px">
          <el-tab-pane
            v-for="sheet in excelStore.sheets"
            :key="sheet.name"
            :label="`${sheet.name} (${sheet.rows}行)`"
            :name="sheet.name"
            lazy
          >
            <SheetPreview :upload-id="excelStore.upload!.upload_id" :sheet-name="sheet.name" />
          </el-tab-pane>
        </el-tabs>
      </div>
    </el-card>

    <!-- Step 2: Clinical Info Form -->
    <el-card v-if="excelStore.upload" shadow="hover" style="margin-bottom: 20px">
      <template #header>
        <div style="display: flex; justify-content: space-between; align-items: center">
          <strong>2. 临床信息</strong>
          <el-select
            v-model="projectType"
            placeholder="项目类型"
            style="width: 250px"
            clearable
            :disabled="isReadOnly"
          >
            <el-option label="结直肠癌301基因+MSI" value="crc_301_msi" />
            <el-option label="结直肠癌358基因+MSI" value="crc_358_msi" />
            <el-option label="MLF基因检测" value="mlf_result" />
            <el-option label="肺癌甲基化" value="lung_methylation" />
          </el-select>
        </div>
      </template>
      <DynamicClinicalForm
        :schema="form.schema.value"
        :form-data="form.formData"
        :errors="form.errors.value"
        :loading="form.loading.value"
        :disabled="isReadOnly"
      />
    </el-card>

    <!-- Step 3: Generate -->
    <el-card v-if="excelStore.upload" shadow="hover">
      <template #header><strong>3. 生成报告</strong></template>
      <el-button
        type="primary"
        size="large"
        :loading="generating"
        :disabled="isReadOnly"
        @click="handleGenerate"
      >
        {{ actionButtonLabel }}
      </el-button>

      <!-- Result -->
      <div v-if="result" style="margin-top: 16px">
        <el-result
          :icon="result.success ? 'success' : 'error'"
          :title="result.success ? '报告生成成功' : '报告生成失败'"
          :sub-title="result.duration_seconds ? `耗时 ${result.duration_seconds.toFixed(1)} 秒` : ''"
        >
          <template #extra>
            <el-button
              v-if="result.success && result.task_id"
              type="primary"
              @click="downloadReport(result.task_id)"
            >
              下载报告
            </el-button>
          </template>
        </el-result>
        <el-alert
          v-for="(err, i) in result.errors"
          :key="i"
          :title="err"
          type="error"
          show-icon
          :closable="false"
          style="margin-bottom: 4px"
        />
        <el-alert
          v-for="(warn, i) in result.warnings"
          :key="'w' + i"
          :title="warn"
          type="warning"
          show-icon
          :closable="false"
          style="margin-bottom: 4px"
        />
      </div>
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { useExcelStore } from '@/stores/excel'
import { useDynamicForm } from '@/composables/useDynamicForm'
import { reportApi, type GenerateResult } from '@/api/report'
import { taskApi, type TaskItem } from '@/api/task'
import DynamicClinicalForm from '@/components/clinical/DynamicClinicalForm.vue'
import SheetPreview from '@/components/excel/SheetPreview.vue'

const excelStore = useExcelStore()
const route = useRoute()
const router = useRouter()

const projectType = ref<string | null>(null)
const generating = ref(false)
const result = ref<GenerateResult | null>(null)
const currentTask = ref<TaskItem | null>(null)
const taskLoading = ref(false)

const taskIdQuery = computed(() => {
  const raw = route.query.task_id
  return typeof raw === 'string' && raw ? raw : null
})

const isReadOnly = computed(() => route.query.mode === 'view')
const pageTitle = computed(() => {
  if (!currentTask.value) return '生成报告'
  return isReadOnly.value ? '查看历史任务' : '编辑并重新生成报告'
})
const actionButtonLabel = computed(() => currentTask.value ? '重新生成报告' : '生成报告')
const taskStatusLabel = computed(() => {
  const status = currentTask.value?.status
  return {
    completed: '已完成',
    failed: '失败',
    running: '运行中',
    pending: '待执行',
    cancelled: '已取消',
  }[status || ''] || '-'
})

// Initialize projectType from detection
watch(
  () => excelStore.upload?.detected_project_type,
  (type) => {
    if (type && !currentTask.value) projectType.value = type
  },
)

// Dynamic form driven by project type
const form = useDynamicForm(projectType)

// Auto-merge Excel values when upload completes
watch(
  () => excelStore.singleValues,
  (vals) => {
    if (vals && Object.keys(vals).length > 0) {
      form.mergeExcelValues(vals)
    }
  },
)

watch(
  taskIdQuery,
  async (taskId) => {
    if (taskId) {
      await loadExistingTask(taskId)
    } else {
      resetTaskContext()
    }
  },
  { immediate: true },
)

function resetTaskContext() {
  const hadTask = !!currentTask.value
  currentTask.value = null
  taskLoading.value = false
  result.value = null
  if (hadTask) {
    excelStore.reset()
    form.reset()
    projectType.value = null
    return
  }
  projectType.value = excelStore.upload?.detected_project_type || null
}

async function waitForFormReady(maxAttempts = 40) {
  for (let i = 0; i < maxAttempts; i += 1) {
    if (!form.loading.value && form.schema.value) return
    await new Promise((resolve) => setTimeout(resolve, 50))
  }
}

async function loadExistingTask(taskId: string) {
  taskLoading.value = true
  result.value = null
  try {
    const task = await taskApi.get(taskId)
    currentTask.value = task
    if (task.upload_id) {
      await excelStore.loadUpload(task.upload_id)
    }

    form.reset()
    projectType.value = task.project_type || excelStore.upload?.detected_project_type || null
    await waitForFormReady()
    form.mergeExcelValues(excelStore.singleValues)
    if (task.clinical_info_snapshot) {
      form.mergePatientInfo(task.clinical_info_snapshot)
    }
  } catch (err: any) {
    ElMessage.error(err.response?.data?.detail || '任务加载失败')
  } finally {
    taskLoading.value = false
  }
}

async function handleFileChange(uploadFile: any) {
  const file = uploadFile.raw || uploadFile
  if (!file) return
  if (taskIdQuery.value) {
    await router.replace({ name: 'generate' })
    currentTask.value = null
  }
  result.value = null
  try {
    await excelStore.uploadFile(file)
    ElMessage.success('Excel 上传成功')
  } catch (err: any) {
    ElMessage.error(err.response?.data?.detail || 'Excel 上传失败')
  }
}

async function handleGenerate() {
  if (!excelStore.upload) return
  if (isReadOnly.value) return
  if (!form.validate()) {
    ElMessage.warning('请填写必填字段')
    return
  }

  generating.value = true
  result.value = null
  try {
    result.value = await reportApi.generate({
      upload_id: excelStore.upload.upload_id,
      clinical_info: form.getCleanValues(),
      project_type: projectType.value,
    })
    if (result.value.success) {
      ElMessage.success('报告生成成功')
      currentTask.value = await taskApi.get(result.value.task_id)
      await router.replace({ name: 'generate', query: { task_id: result.value.task_id, mode: 'edit' } })
    } else {
      ElMessage.error('报告生成失败')
    }
  } catch (err: any) {
    ElMessage.error(err.response?.data?.error || '报告生成异常')
  } finally {
    generating.value = false
  }
}

async function downloadReport(taskId: string) {
  try {
    await reportApi.download(taskId)
  } catch (err: any) {
    ElMessage.error(err.response?.data?.detail || '下载失败，请重新登录后再试')
  }
}

function switchToEdit() {
  if (!taskIdQuery.value) return
  router.replace({ name: 'generate', query: { task_id: taskIdQuery.value, mode: 'edit' } })
}
</script>
