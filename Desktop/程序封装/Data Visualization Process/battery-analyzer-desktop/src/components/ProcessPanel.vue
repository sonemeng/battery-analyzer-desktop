<template>
  <div class="process-panel">
    <NCard title="数据处理" style="height: 100%">
      <!-- 文件夹选择区域 -->
      <div class="folder-section">
        <NSpace align="center" style="margin-bottom: 16px;">
          <NButton @click="selectFolder" type="primary" :loading="isSelecting">
            选择文件夹
          </NButton>
          <NText v-if="currentFolder" class="folder-path">{{ currentFolder }}</NText>
        </NSpace>

        <NAlert v-if="!currentFolder" type="warning" style="margin-bottom: 16px;">
          请先在配置管理中设置数据文件夹路径，或点击上方按钮选择文件夹
        </NAlert>
      </div>

      <!-- 文件列表 -->
      <div v-if="files.length > 0" class="files-section">
        <NText strong style="margin-bottom: 8px; display: block;">
          文件列表 ({{ files.length }} 个文件)
        </NText>
        <div class="file-list">
          <div v-for="file in files.slice(0, 5)" :key="file" class="file-item">
            <NIcon :component="DocumentIcon" style="margin-right: 8px;" />
            <NText>{{ file }}</NText>
          </div>
          <NText v-if="files.length > 5" depth="3" style="margin-top: 8px;">
            ... 还有 {{ files.length - 5 }} 个文件
          </NText>
        </div>
      </div>

      <!-- 处理控制 -->
      <div class="process-controls" style="margin-top: 24px;">
        <NSpace>
          <NButton
            @click="startProcessing"
            type="primary"
            :loading="isProcessing"
            :disabled="!currentFolder || files.length === 0"
          >
            开始处理
          </NButton>
          <NButton @click="clearLogs" :disabled="isProcessing">
            清空日志
          </NButton>
        </NSpace>
      </div>

      <!-- 进度条 -->
      <div v-if="isProcessing" class="progress-section" style="margin-top: 16px;">
        <NProgress
          type="line"
          :percentage="progress"
          :show-indicator="true"
          status="info"
        />
        <NText style="margin-top: 8px;">处理进度: {{ progress }}%</NText>
      </div>

      <!-- 日志输出 -->
      <div class="logs-section" style="margin-top: 24px;">
        <NText strong style="margin-bottom: 8px; display: block;">处理日志</NText>
        <div class="log-container">
          <div v-for="(log, index) in logs" :key="index" class="log-entry">
            <NText :type="log.type === 'error' ? 'error' : log.type === 'success' ? 'success' : 'default'">
              {{ log.timestamp }} {{ log.message }}
            </NText>
          </div>
        </div>
      </div>
    </NCard>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import {
  NCard, NText, NButton, NSpace, NAlert, NIcon, NProgress,
  useMessage
} from 'naive-ui'
import { Document as DocumentIcon } from '@vicons/ionicons5'
import { invoke } from '@tauri-apps/api/core'
import { open } from '@tauri-apps/plugin-dialog'
import { useConfigStore } from '../stores/config'

const configStore = useConfigStore()
const message = useMessage()
// const dialog = useDialog() // 暂时不需要

// 响应式数据
const selectedFolder = ref('')
const files = ref<string[]>([])
const isSelecting = ref(false)
const isProcessing = ref(false)
const progress = ref(0)
const logs = ref<Array<{timestamp: string, message: string, type: string}>>([])

// 计算属性：当前使用的文件夹路径
const currentFolder = computed(() => {
  return selectedFolder.value || configStore.config.input_folder
})

// 添加日志
const addLog = (message: string, type: 'info' | 'success' | 'error' = 'info') => {
  const timestamp = new Date().toLocaleTimeString()
  logs.value.push({ timestamp, message, type })
}

// 选择文件夹
const selectFolder = async () => {
  try {
    isSelecting.value = true
    const selected = await open({
      directory: true,
      multiple: false,
      title: '选择数据文件夹'
    })

    if (selected && typeof selected === 'string') {
      selectedFolder.value = selected
      await loadFiles()
      addLog(`已选择文件夹: ${selected}`, 'success')
    }
  } catch (error) {
    console.error('选择文件夹失败:', error)
    message.error('选择文件夹失败')
  } finally {
    isSelecting.value = false
  }
}

// 加载文件列表
const loadFiles = async () => {
  if (!currentFolder.value) return

  try {
    const fileList = await invoke<string[]>('list_excel_files', {
      folderPath: currentFolder.value
    })
    files.value = fileList
    addLog(`找到 ${fileList.length} 个Excel文件`, 'info')
  } catch (error) {
    console.error('加载文件列表失败:', error)
    message.error('加载文件列表失败')
    files.value = []
  }
}

// 开始处理
const startProcessing = async () => {
  if (!currentFolder.value) {
    message.warning('请先选择文件夹')
    return
  }

  if (files.value.length === 0) {
    message.warning('没有找到可处理的文件')
    return
  }

  try {
    isProcessing.value = true
    progress.value = 0

    addLog(`开始处理 ${files.value.length} 个文件`, 'info')
    addLog('调用Python脚本处理数据...', 'info')

    // 模拟进度更新
    const progressInterval = setInterval(() => {
      if (progress.value < 90) {
        progress.value += 10
      }
    }, 500)

    // 获取最新的完整配置
    const latestConfig = configStore.toBackendConfig()

    // 调用后端处理（使用最新配置）
    const result = await invoke<string>('process_battery_data', {
      config: {
        ...latestConfig,  // 使用完整的最新配置
        input_folder: currentFolder.value,
        output_folder: latestConfig.output_folder || currentFolder.value
      }
    })

    clearInterval(progressInterval)
    progress.value = 100

    addLog('数据处理完成！', 'success')
    addLog(result, 'info')
    message.success('数据处理完成！')

  } catch (error) {
    console.error('处理失败:', error)
    addLog(`处理失败: ${error}`, 'error')
    message.error('处理失败')
  } finally {
    isProcessing.value = false
  }
}

// 清空日志
const clearLogs = () => {
  logs.value = []
}

// 组件挂载时加载文件
onMounted(() => {
  if (currentFolder.value) {
    loadFiles()
  }
})
</script>

<style scoped>
.process-panel {
  height: 100%;
}

.folder-path {
  color: #18a058;
  font-weight: 500;
}

.file-list {
  max-height: 200px;
  overflow-y: auto;
  border: 1px solid #e0e0e6;
  border-radius: 6px;
  padding: 12px;
  background-color: #fafafa;
}

.file-item {
  display: flex;
  align-items: center;
  padding: 4px 0;
  border-bottom: 1px solid #f0f0f0;
}

.file-item:last-child {
  border-bottom: none;
}

.log-container {
  max-height: 300px;
  overflow-y: auto;
  border: 1px solid #e0e0e6;
  border-radius: 6px;
  padding: 12px;
  background-color: #fafafa;
  font-family: 'Consolas', 'Monaco', monospace;
  font-size: 12px;
}

.log-entry {
  margin-bottom: 4px;
  line-height: 1.4;
}
</style>
