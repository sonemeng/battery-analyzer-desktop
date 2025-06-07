<script setup lang="ts">
import { ref, h, onMounted, onUnmounted } from 'vue'
import { invoke } from '@tauri-apps/api/core'
import { listen, UnlistenFn } from '@tauri-apps/api/event'
import { useConfigStore } from './stores/config'
import {
  NConfigProvider,
  NLayout,
  NLayoutHeader,
  NLayoutContent,
  NLayoutSider,
  NMenu,
  NSpace,
  NButton,
  NIcon,
  NText,
  NCard,
  NTabs,
  NTabPane,
  NFormItem,
  NInput,
  NInputNumber,
  NCheckbox,
  NCheckboxGroup,
  NRadio,
  NRadioGroup,
  NSelect,
  NAlert,
  NGrid,
  NGridItem,
  NProgress,
  NScrollbar,
  NTag,
  NTime,
  NEmpty,
  NList,
  NListItem,
  NThing,
  darkTheme,
  zhCN,
  dateZhCN
} from 'naive-ui'
import {
  SettingsOutline,
  PlayOutline,
  DocumentTextOutline,
  BarChartOutline,
  MoonOutline,
  SunnyOutline,
  FolderOpenOutline,
  DocumentOutline,
  RefreshOutline,
  CheckmarkCircleOutline,
  CloseCircleOutline,
  TimeOutline,
  TrashOutline,
  PauseOutline,
  StopOutline
} from '@vicons/ionicons5'

// 当前活动面板
const activePanel = ref('config')

// 主题设置
const isDark = ref(false)

// 配置store
const configStore = useConfigStore()

// 配置相关的响应式变量已移除，直接使用configStore

// 文件信息接口（与 Rust 后端保持一致）
interface FileInfo {
  name: string
  path: string
  size: number
  last_modified: number
  is_excel: boolean
}

// 前端文件状态接口
interface FileWithStatus extends FileInfo {
  status: 'pending' | 'processing' | 'completed' | 'error'
  lastModified: Date
}

// 数据处理相关状态
const selectedFolder = ref('')
const fileList = ref<FileWithStatus[]>([])
const isProcessing = ref(false)
const processingProgress = ref(0)
const processingStatus = ref('idle') // 'idle' | 'running' | 'paused' | 'completed' | 'error'
const processingLogs = ref<Array<{
  timestamp: Date
  level: 'info' | 'warning' | 'error' | 'success'
  message: string
}>>([])
const processedCount = ref(0)
const totalCount = ref(0)

// 事件监听器
let progressUnlisten: UnlistenFn | null = null

// 菜单选项
const menuOptions = [
  {
    label: '配置管理',
    key: 'config',
    icon: () => h(NIcon, null, { default: () => h(SettingsOutline) })
  },
  {
    label: '数据处理',
    key: 'process',
    icon: () => h(NIcon, null, { default: () => h(PlayOutline) })
  },
  {
    label: '结果查看',
    key: 'results',
    icon: () => h(NIcon, null, { default: () => h(BarChartOutline) })
  },
  {
    label: '文档帮助',
    key: 'docs',
    icon: () => h(NIcon, null, { default: () => h(DocumentTextOutline) })
  }
]

// 处理菜单选择
const handleMenuSelect = (key: string) => {
  activePanel.value = key
}

// 数据处理相关方法
const selectFolder = async () => {
  try {
    // 临时解决方案：使用 prompt 让用户输入路径
    // TODO: 修复 Tauri 对话框 API 后替换为文件夹选择对话框
    const folderPath = prompt('请输入包含电池数据文件的文件夹路径:', 'C:\\Users\\XIGU\\Desktop\\程序封装\\Data Visualization Process')

    console.log('输入的文件夹路径:', folderPath)

    if (folderPath && folderPath.trim() !== '') {
      selectedFolder.value = folderPath.trim()
      addLog('success', `已设置文件夹: ${folderPath}`)
      await loadFileList()
    } else {
      addLog('info', '用户取消了文件夹选择')
    }
  } catch (error) {
    console.error('选择文件夹失败:', error)
    addLog('error', `选择文件夹失败: ${String(error)}`)
  }
}

const loadFileList = async () => {
  if (!selectedFolder.value) {
    addLog('warning', '请先选择文件夹')
    return
  }

  try {
    addLog('info', '正在读取文件列表...')

    // 调用 Rust 后端命令读取目录
    const files = await invoke<FileInfo[]>('read_directory', {
      path: selectedFolder.value
    })

    // 转换为前端格式
    fileList.value = files.map(file => ({
      ...file,
      lastModified: new Date(file.last_modified * 1000), // 转换时间戳
      status: 'pending' as const
    }))

    totalCount.value = fileList.value.length
    addLog('success', `找到 ${totalCount.value} 个Excel文件`)

  } catch (error) {
    console.error('读取文件列表失败:', error)
    addLog('error', `读取文件列表失败: ${error}`)
    fileList.value = []
    totalCount.value = 0
  }
}

const addLog = (level: 'info' | 'warning' | 'error' | 'success', message: string) => {
  processingLogs.value.push({
    timestamp: new Date(),
    level,
    message
  })
}

const startProcessing = async () => {
  if (fileList.value.length === 0) {
    addLog('error', '没有找到可处理的文件')
    return
  }

  if (!selectedFolder.value) {
    addLog('error', '请先选择文件夹')
    return
  }

  isProcessing.value = true
  processingStatus.value = 'running'
  processedCount.value = 0
  processingProgress.value = 0

  try {
    // 构建处理配置 - 使用统一的配置格式
    const backendConfig = configStore.toBackendConfig()
    const processConfig = {
      ...backendConfig,
      input_folder: selectedFolder.value,
      output_folder: backendConfig.output_folder || selectedFolder.value
      // 其他配置已经在backendConfig中包含
    }

    // 调用Rust后端处理数据（进度和日志现在通过事件实时更新）
    await invoke<string>('process_battery_data', { config: processConfig })

    // 更新所有文件状态为已完成
    fileList.value.forEach(file => {
      file.status = 'completed'
    })

    processedCount.value = totalCount.value

  } catch (error) {
    processingStatus.value = 'error'
    addLog('error', `处理失败: ${String(error)}`)
    console.error('处理错误:', error)

    // 更新文件状态为错误
    fileList.value.forEach(file => {
      if (file.status === 'processing') {
        file.status = 'error'
      }
    })
  } finally {
    isProcessing.value = false
  }
}

const pauseProcessing = () => {
  processingStatus.value = 'paused'
  addLog('warning', '处理已暂停')
}

const stopProcessing = () => {
  processingStatus.value = 'idle'
  isProcessing.value = false
  fileList.value.forEach(file => {
    if (file.status === 'processing') {
      file.status = 'pending'
    }
  })
  addLog('warning', '处理已停止')
}

const clearLogs = () => {
  processingLogs.value = []
}

const formatFileSize = (bytes: number) => {
  if (bytes === 0) return '0 B'
  const k = 1024
  const sizes = ['B', 'KB', 'MB', 'GB']
  const i = Math.floor(Math.log(bytes) / Math.log(k))
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]
}

// 界面辅助方法
const getFileStatusColor = (status: string) => {
  switch (status) {
    case 'pending': return '#909399'
    case 'processing': return '#409EFF'
    case 'completed': return '#67C23A'
    case 'error': return '#F56C6C'
    default: return '#909399'
  }
}

const getFileStatusType = (status: string) => {
  switch (status) {
    case 'pending': return 'default'
    case 'processing': return 'info'
    case 'completed': return 'success'
    case 'error': return 'error'
    default: return 'default'
  }
}

const getFileStatusText = (status: string) => {
  switch (status) {
    case 'pending': return '待处理'
    case 'processing': return '处理中'
    case 'completed': return '已完成'
    case 'error': return '处理失败'
    default: return '未知'
  }
}

const getProgressStatus = () => {
  switch (processingStatus.value) {
    case 'running': return 'default'
    case 'completed': return 'success'
    case 'error': return 'error'
    case 'paused': return 'warning'
    default: return 'default'
  }
}

const getStatusAlertType = () => {
  switch (processingStatus.value) {
    case 'running': return 'info'
    case 'completed': return 'success'
    case 'error': return 'error'
    case 'paused': return 'warning'
    default: return 'info'
  }
}

const getStatusText = () => {
  switch (processingStatus.value) {
    case 'running': return '正在处理文件...'
    case 'completed': return '所有文件处理完成！'
    case 'error': return '处理过程中出现错误'
    case 'paused': return '处理已暂停'
    default: return ''
  }
}

const getLogTagType = (level: string) => {
  switch (level) {
    case 'info': return 'info'
    case 'warning': return 'warning'
    case 'error': return 'error'
    case 'success': return 'success'
    default: return 'default'
  }
}

const getLogLevelText = (level: string) => {
  switch (level) {
    case 'info': return '信息'
    case 'warning': return '警告'
    case 'error': return '错误'
    case 'success': return '成功'
    default: return '未知'
  }
}

// 设置进度事件监听器
onMounted(async () => {
  try {
    progressUnlisten = await listen('processing-progress', (event) => {
      const payload = event.payload as { progress: number; message: string; level: string }

      // 更新进度条
      processingProgress.value = payload.progress

      // 添加日志
      addLog(payload.level as any, payload.message)

      // 更新处理状态
      if (payload.progress >= 100) {
        processingStatus.value = 'completed'
        isProcessing.value = false
      } else if (payload.level === 'error') {
        processingStatus.value = 'error'
        isProcessing.value = false
      } else if (payload.progress > 0) {
        processingStatus.value = 'running'
      }
    })
  } catch (error) {
    console.error('设置进度监听器失败:', error)
  }
})

// 清理事件监听器
onUnmounted(() => {
  if (progressUnlisten) {
    progressUnlisten()
  }
})
</script>

<template>
  <NConfigProvider
    :theme="isDark ? darkTheme : null"
    :locale="zhCN"
    :date-locale="dateZhCN"
  >
    <NLayout style="height: 100vh">
      <!-- 顶部标题栏 -->
      <NLayoutHeader style="height: 64px; padding: 0 24px" bordered>
        <div style="display: flex; align-items: center; height: 100%">
          <NIcon size="32" color="#1890ff" style="margin-right: 12px">
            <BarChartOutline />
          </NIcon>
          <NText style="font-size: 20px; font-weight: 600">
            电池数据分析器
          </NText>
          <div style="flex: 1"></div>
          <NSpace>
            <NButton
              quaternary
              circle
              @click="isDark = !isDark"
            >
              <template #icon>
                <NIcon>
                  <MoonOutline v-if="!isDark" />
                  <SunnyOutline v-else />
                </NIcon>
              </template>
            </NButton>
          </NSpace>
        </div>
      </NLayoutHeader>

      <NLayout has-sider>
        <!-- 左侧导航菜单 -->
        <NLayoutSider
          bordered
          collapse-mode="width"
          :collapsed-width="64"
          :width="240"
          show-trigger
        >
          <NMenu
            :collapsed-width="64"
            :collapsed-icon-size="22"
            :options="menuOptions"
            :value="activePanel"
            @update:value="handleMenuSelect"
          />
        </NLayoutSider>

        <!-- 主内容区域 -->
        <NLayoutContent style="padding: 24px">
          <div style="height: 100%">
            <!-- 配置管理面板 -->
            <div v-if="activePanel === 'config'" style="height: 100%; overflow-y: auto;">
              <!-- 页面标题 -->
              <div style="margin-bottom: 24px;">
                <NText style="font-size: 24px; font-weight: 600;">配置管理</NText>
                <NText depth="3" style="display: block; margin-top: 8px;">
                  专业的电池数据分析配置界面，支持80+个参数的精细调节
                </NText>
              </div>

              <!-- 配置分类标签页 -->
              <NTabs type="line" animated>
                <!-- 基础配置 -->
                <NTabPane name="basic" tab="基础配置" display-directive="show:lazy">
                  <NSpace vertical size="large">
                    <!-- 数据文件配置 -->
                    <NCard title="数据文件配置">
                      <NSpace vertical size="medium">
                        <NAlert type="info" style="margin: 16px 0;">
                          <template #header>输出说明</template>
                          程序生成的汇总表和可视化文件将自动保存在数据文件夹中，无需单独指定输出路径。数据文件夹在"数据处理"页面选择。
                        </NAlert>
                        <NFormItem label="输出文件配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">汇总表文件名前缀:</NText>
                                <NInput
                                  placeholder="LIMS数据汇总表"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <NCheckbox :default-checked="true">自动添加时间戳</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <NCheckbox :default-checked="true">生成可视化文件夹</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="文件名过滤">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <NCheckbox :default-checked="true">排除以'_1.xlsx'结尾的文件</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <NCheckbox :default-checked="true">排除隐藏文件(~开头)</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <NCheckbox :default-checked="true">只处理Excel文件(.xlsx)</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- Excel读取配置 -->
                    <NCard title="Excel读取配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="Excel读取引擎">
                          <div class="config-row">
                            <NText class="config-label">读取引擎:</NText>
                            <!-- 绑定到 excel_engine 参数 -->
                            <NSelect
                              v-model:value="configStore.config.excel_engine"
                              class="config-input"
                              :options="[
                                { label: 'Calamine (推荐)', value: 'calamine' },
                                { label: 'OpenPyXL', value: 'openpyxl' }
                              ]"
                            />
                          </div>
                        </NFormItem>
                        <NFormItem label="工作表名称">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">循环数据工作表:</NText>
                                <!-- 绑定到 cycle_sheet_name 参数 -->
                                <NInput
                                  v-model:value="configStore.config.cycle_sheet_name"
                                  placeholder="Cycle"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">测试数据工作表:</NText>
                                <!-- 绑定到 test_sheet_name 参数 -->
                                <NInput
                                  v-model:value="configStore.config.test_sheet_name"
                                  placeholder="test"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 程序运行配置 -->
                    <NCard title="程序运行配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="运行模式">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <!-- 绑定到 verbose 参数 -->
                              <NCheckbox v-model:checked="configStore.config.verbose">显示详细输出</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 enable_progress_bar 参数 -->
                              <NCheckbox v-model:checked="configStore.config.enable_progress_bar">显示进度条</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 backup_original_data 参数 -->
                              <NCheckbox v-model:checked="configStore.config.backup_original_data">备份原始数据</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 auto_open_results 参数 -->
                              <NCheckbox v-model:checked="configStore.config.auto_open_results">自动打开结果文件</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="性能配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">文件处理分块大小:</NText>
                                <!-- 绑定到 chunk_size 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.chunk_size"
                                  :min="10"
                                  :max="200"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">内存使用限制(MB):</NText>
                                <!-- 绑定到 memory_limit_mb 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.memory_limit_mb"
                                  :min="100"
                                  :max="2000"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="日志级别">
                          <div class="config-row">
                            <NText class="config-label">日志级别:</NText>
                            <!-- 绑定到 log_level 参数 -->
                            <NSelect
                              v-model:value="configStore.config.log_level"
                              class="config-input"
                              :options="[
                                { label: 'DEBUG', value: 'DEBUG' },
                                { label: 'INFO', value: 'INFO' },
                                { label: 'WARNING', value: 'WARNING' },
                                { label: 'ERROR', value: 'ERROR' }
                              ]"
                            />
                          </div>
                        </NFormItem>
                      </NSpace>
                    </NCard>
                  </NSpace>
                </NTabPane>

                <!-- 异常检测配置 -->
                <NTabPane name="outlier" tab="异常检测配置" display-directive="show:lazy">
                  <NSpace vertical size="large">
                    <!-- 异常检测方法 -->
                    <NCard title="异常检测方法">
                      <NSpace vertical size="medium">
                        <NAlert type="info" style="margin-bottom: 16px;">
                          <template #header>方法说明</template>
                          程序支持两种异常检测方法：改良箱线图法（默认）和Z-score+MAD法。每次只能选择一种方法进行异常检测。
                        </NAlert>
                        <NFormItem label="检测方法">
                          <!-- 绑定到 outlier_method 参数 -->
                          <NRadioGroup v-model:value="configStore.config.outlier_method">
                            <NSpace>
                              <NRadio value="boxplot">改良箱线图法</NRadio>
                              <NRadio value="zscore_mad">Z-score+MAD法</NRadio>
                            </NSpace>
                          </NRadioGroup>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 改良箱线图配置 -->
                    <NCard title="改良箱线图配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="极差阈值">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首放_极差阈值:</NText>
                                <!-- 绑定到 boxplot_threshold_discharge 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.boxplot_threshold_discharge"
                                  :min="1"
                                  :max="50"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效_极差阈值:</NText>
                                <!-- 绑定到 boxplot_threshold_efficiency 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.boxplot_threshold_efficiency"
                                  :min="0.1"
                                  :max="10"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="迭代参数">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最大迭代次数:</NText>
                                <!-- 绑定到 max_iterations 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.max_iterations"
                                  :min="1"
                                  :max="20"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">箱线图收缩因子:</NText>
                                <!-- 绑定到 boxplot_shrink_factor 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.boxplot_shrink_factor"
                                  :min="0.8"
                                  :max="1.0"
                                  :step="0.01"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- Z-score+MAD配置 -->
                    <NCard title="Z-score+MAD配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="Z-score阈值设置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首放容量阈值:</NText>
                                <!-- 绑定到 zscore_threshold_discharge 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_threshold_discharge"
                                  :min="1"
                                  :max="5"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效阈值:</NText>
                                <!-- 绑定到 zscore_threshold_efficiency 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_threshold_efficiency"
                                  :min="1"
                                  :max="5"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首圈电压阈值:</NText>
                                <!-- 绑定到 zscore_threshold_voltage 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_threshold_voltage"
                                  :min="1"
                                  :max="5"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首圈能量阈值:</NText>
                                <!-- 绑定到 zscore_threshold_energy 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_threshold_energy"
                                  :min="1"
                                  :max="5"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="MAD配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">MAD常数:</NText>
                                <!-- 绑定到 zscore_mad_constant 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_mad_constant"
                                  :min="0.1"
                                  :max="2"
                                  :step="0.0001"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">MAD最小值比例:</NText>
                                <!-- 绑定到 zscore_min_mad_ratio 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_min_mad_ratio"
                                  :min="0.001"
                                  :max="1"
                                  :step="0.001"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 zscore_use_time_series 参数 -->
                              <NCheckbox v-model:checked="configStore.config.zscore_use_time_series">使用时间序列分解</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="时间序列分解">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">STL最小样本数:</NText>
                                <!-- 绑定到 zscore_min_samples_for_stl 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.zscore_min_samples_for_stl"
                                  :min="5"
                                  :max="50"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 zscore_generate_plots 参数 -->
                              <NCheckbox v-model:checked="configStore.config.zscore_generate_plots">生成分布图</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>
                  </NSpace>
                </NTabPane>

                <!-- 参考通道选择 -->
                <NTabPane name="reference" tab="参考通道选择" display-directive="show:lazy">
                  <NSpace vertical size="large">
                    <!-- 参考通道选择方法 -->
                    <NCard title="参考通道选择方法">
                      <NSpace vertical size="medium">
                        <NAlert type="info" style="margin-bottom: 16px;">
                          <template #header>方法说明</template>
                          程序支持多种参考通道选择方法，按优先级顺序尝试。
                        </NAlert>
                        <NFormItem label="选择方法">
                          <!-- 绑定到 reference_channel_method 参数 -->
                          <NRadioGroup v-model:value="configStore.config.reference_channel_method">
                            <NSpace>
                              <NRadio value="retention_curve_mse">容量保持率曲线比较</NRadio>
                              <NRadio value="pca">PCA方法</NRadio>
                              <NRadio value="traditional">传统方法</NRadio>
                            </NSpace>
                          </NRadioGroup>
                        </NFormItem>
                        <NFormItem label="方法优先级">
                          <NAlert type="info" style="margin-bottom: 8px;">
                            程序按以下优先级顺序尝试各种方法
                          </NAlert>
                          <NSpace vertical>
                            <NText>1. 容量保持率曲线比较</NText>
                            <NText>2. PCA多特征分析</NText>
                            <NText>3. 传统方法(首放接近均值)</NText>
                          </NSpace>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 容量保持率配置 -->
                    <NCard title="容量保持率曲线比较配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="基本参数">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <!-- 绑定到 capacity_retention_enabled 参数 -->
                              <NCheckbox v-model:checked="configStore.config.capacity_retention_enabled">启用容量保持率曲线比较</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最小通道数:</NText>
                                <!-- 绑定到 capacity_retention_min_channels 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_retention_min_channels"
                                  :min="1"
                                  :max="20"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最小循环次数:</NText>
                                <!-- 绑定到 capacity_retention_min_cycles 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_retention_min_cycles"
                                  :min="1"
                                  :max="50"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最大循环次数:</NText>
                                <!-- 绑定到 capacity_retention_max_cycles 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_retention_max_cycles"
                                  :min="100"
                                  :max="2000"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">循环步长:</NText>
                                <!-- 绑定到 capacity_retention_cycle_step 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_retention_cycle_step"
                                  :min="1"
                                  :max="10"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 capacity_retention_use_raw_capacity 参数 -->
                              <NCheckbox v-model:checked="configStore.config.capacity_retention_use_raw_capacity">使用原始容量数据</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 pca_visualization_enabled 参数 -->
                              <NCheckbox v-model:checked="configStore.config.pca_visualization_enabled">启用可视化</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="权重配置">
                          <NSpace vertical>
                            <!-- 绑定到 capacity_retention_use_weighted_mse 参数 -->
                            <NCheckbox v-model:checked="configStore.config.capacity_retention_use_weighted_mse">使用加权MSE</NCheckbox>
                            <NSpace>
                              <NText>权重方法:</NText>
                              <!-- 绑定到 capacity_retention_weight_method 参数 -->
                              <NSelect v-model:value="configStore.config.capacity_retention_weight_method" :options="[
                                { label: '线性增长', value: 'linear' },
                                { label: '指数增长', value: 'exp' },
                                { label: '恒定权重', value: 'constant' }
                              ]" />
                            </NSpace>
                            <NSpace>
                              <NText>权重因子:</NText>
                              <!-- 绑定到 capacity_retention_weight_factor 参数 -->
                              <NInputNumber v-model:value="configStore.config.capacity_retention_weight_factor" :min="0.1" :max="5" :step="0.1" />
                            </NSpace>
                            <NSpace>
                              <NText>后期循环权重倍数:</NText>
                              <!-- 绑定到 capacity_retention_late_cycles_emphasis 参数 -->
                              <NInputNumber v-model:value="configStore.config.capacity_retention_late_cycles_emphasis" :min="1.0" :max="5.0" :step="0.1" />
                            </NSpace>
                            <!-- 绑定到 capacity_retention_dynamic_range 参数 -->
                            <NCheckbox v-model:checked="configStore.config.capacity_retention_dynamic_range">动态确定循环范围</NCheckbox>
                          </NSpace>
                        </NFormItem>
                        <NFormItem label="插值配置">
                          <NSpace vertical>
                            <NSpace>
                              <NText>插值方法:</NText>
                              <!-- 绑定到 capacity_retention_interpolation_method 参数 -->
                              <NSelect v-model:value="configStore.config.capacity_retention_interpolation_method" :options="[
                                { label: '线性插值', value: 'linear' },
                                { label: '三次插值', value: 'cubic' },
                                { label: '最近邻插值', value: 'nearest' }
                              ]" />
                            </NSpace>
                          </NSpace>
                        </NFormItem>
                        <NFormItem label="多维度权重配置">
                          <NSpace vertical>
                            <!-- 绑定到 capacity_retention_include_voltage 参数 -->
                            <NCheckbox v-model:checked="configStore.config.capacity_retention_include_voltage">包含电压保持率</NCheckbox>
                            <!-- 绑定到 capacity_retention_include_energy 参数 -->
                            <NCheckbox v-model:checked="configStore.config.capacity_retention_include_energy">包含能量保持率</NCheckbox>
                            <NSpace>
                              <NText>容量保持率权重:</NText>
                              <!-- 绑定到 capacity_retention_capacity_weight 参数 -->
                              <NInputNumber v-model:value="configStore.config.capacity_retention_capacity_weight" :min="0.0" :max="1.0" :step="0.1" />
                            </NSpace>
                            <NSpace>
                              <NText>电压保持率权重:</NText>
                              <!-- 绑定到 capacity_retention_voltage_weight 参数 -->
                              <NInputNumber v-model:value="configStore.config.capacity_retention_voltage_weight" :min="0.0" :max="1.0" :step="0.1" />
                            </NSpace>
                            <NSpace>
                              <NText>能量保持率权重:</NText>
                              <!-- 绑定到 capacity_retention_energy_weight 参数 -->
                              <NInputNumber v-model:value="configStore.config.capacity_retention_energy_weight" :min="0.0" :max="1.0" :step="0.1" />
                            </NSpace>
                            <NSpace>
                              <NText>电压保持率列名:</NText>
                              <!-- 绑定到 capacity_retention_voltage_column 参数 -->
                              <NInput v-model:value="configStore.config.capacity_retention_voltage_column" />
                            </NSpace>
                            <NSpace>
                              <NText>能量保持率列名:</NText>
                              <!-- 绑定到 capacity_retention_energy_column 参数 -->
                              <NInput v-model:value="configStore.config.capacity_retention_energy_column" />
                            </NSpace>
                          </NSpace>
                        </NFormItem>
                        <NFormItem label="容量保持率列">
                          <NAlert type="info" style="margin-bottom: 8px;">
                            <template #header>容量保持率分析方法</template>
                            程序支持多种容量保持率分析方法，会根据数据自动选择最佳方法
                          </NAlert>
                          <NSpace vertical>
                            <NText><strong>主要方法：</strong></NText>
                            <NText depth="3">• 1C首圈到当前循环圈数（推荐，动态识别1C首圈）</NText>
                            <NText><strong>备选方法：</strong></NText>
                            <NText depth="3">• 当前容量保持（首圈到当前圈数）</NText>
                            <NText depth="3">• 100容量保持（首圈到第100圈）</NText>
                            <NText depth="3">• 200容量保持（首圈到第200圈）</NText>
                            <NText><strong>说明：</strong></NText>
                            <NText depth="3">程序会自动识别1C模式并动态确定1C首圈位置，提供更准确的容量保持率分析</NText>
                          </NSpace>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- PCA配置 -->
                    <NCard title="PCA分析配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="PCA分析指标">
                          <!-- 绑定到 pca_default_features 参数 -->
                          <NCheckboxGroup v-model:value="configStore.config.pca_default_features">
                            <NSpace vertical>
                              <NCheckbox value="首放">首放容量</NCheckbox>
                              <NCheckbox value="首圈电压">首圈电压</NCheckbox>
                              <NCheckbox value="Cycle4">Cycle4容量保持率</NCheckbox>
                              <NCheckbox value="首效">首效</NCheckbox>
                              <NCheckbox value="首圈能量">首圈能量</NCheckbox>
                            </NSpace>
                          </NCheckboxGroup>
                        </NFormItem>
                        <NFormItem label="特征阈值配置">
                          <NSpace vertical>
                            <NSpace>
                              <NText>0.1C首充容量范围:</NText>
                              <NInputNumber :min="200" :max="300" :default-value="250" />
                              <NText>-</NText>
                              <NInputNumber :min="300" :max="400" :default-value="350" />
                              <NText>mAh/g</NText>
                            </NSpace>
                            <NSpace>
                              <NText>0.1C首放容量范围:</NText>
                              <NInputNumber :min="150" :max="250" :default-value="200" />
                              <NText>-</NText>
                              <NInputNumber :min="250" :max="350" :default-value="300" />
                              <NText>mAh/g</NText>
                            </NSpace>
                            <NSpace>
                              <NText>首圈电压范围:</NText>
                              <NInputNumber :min="3.0" :max="3.5" :step="0.1" :default-value="3.4" />
                              <NText>-</NText>
                              <NInputNumber :min="3.5" :max="4.0" :step="0.1" :default-value="3.9" />
                              <NText>V</NText>
                            </NSpace>
                            <NSpace>
                              <NText>首圈能量范围:</NText>
                              <NInputNumber :min="500" :max="700" :default-value="600" />
                              <NText>-</NText>
                              <NInputNumber :min="1200" :max="1500" :default-value="1300" />
                              <NText>mWh/g</NText>
                            </NSpace>
                            <NSpace>
                              <NText>首充截止电压上限:</NText>
                              <NInputNumber :min="4.5" :max="5.0" :step="0.01" :default-value="4.7" />
                              <NText>V</NText>
                            </NSpace>
                            <NSpace>
                              <NText>Cycle4容量保持率下限:</NText>
                              <NInputNumber :min="70" :max="90" :default-value="80" />
                              <NText>%</NText>
                            </NSpace>
                            <NSpace>
                              <NText>安全电压阈值:</NText>
                              <NInputNumber :min="4.0" :max="5.0" :step="0.01" :default-value="4.65" />
                              <NText>V</NText>
                            </NSpace>
                          </NSpace>
                        </NFormItem>
                        <NFormItem label="PCA配置">
                          <NSpace vertical>
                            <NSpace>
                              <NText>主成分数量:</NText>
                              <NInputNumber :min="1" :max="10" :default-value="2" />
                            </NSpace>
                            <NCheckbox :default-checked="true">启用可视化</NCheckbox>
                          </NSpace>
                        </NFormItem>
                      </NSpace>
                    </NCard>
                  </NSpace>
                </NTabPane>

                <!-- 阈值配置 -->
                <NTabPane name="thresholds" tab="阈值配置" display-directive="show:lazy">
                  <NSpace vertical size="large">
                    <!-- 数据质量阈值 -->
                    <NCard title="数据质量阈值">
                      <NSpace vertical size="medium">
                        <NFormItem label="异常数据阈值">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首充容量上限(mAh/g):</NText>
                                <!-- 绑定到 abnormal_high_charge 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.abnormal_high_charge"
                                  :min="100"
                                  :max="500"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首充容量下限(mAh/g):</NText>
                                <!-- 绑定到 abnormal_low_charge 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.abnormal_low_charge"
                                  :min="50"
                                  :max="300"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首放容量下限(mAh/g):</NText>
                                <!-- 绑定到 abnormal_low_discharge 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.abnormal_low_discharge"
                                  :min="50"
                                  :max="300"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="1C阈值配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">1C首圈识别比值阈值:</NText>
                                <!-- 绑定到 ratio_threshold 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.ratio_threshold"
                                  :min="0.5"
                                  :max="1.0"
                                  :step="0.01"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">1C首圈识别差值阈值(mAh/g):</NText>
                                <!-- 绑定到 discharge_diff_threshold 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.discharge_diff_threshold"
                                  :min="10"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">过充电压阈值(V):</NText>
                                <!-- 绑定到 overcharge_threshold 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.overcharge_threshold"
                                  :min="4.0"
                                  :max="5.0"
                                  :step="0.01"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效过低阈值(%):</NText>
                                <!-- 绑定到 very_low_efficiency_threshold 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.very_low_efficiency_threshold"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效低阈值(%):</NText>
                                <!-- 绑定到 low_efficiency_threshold 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.low_efficiency_threshold"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="过充风险阈值">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首充截止电压_警告(V):</NText>
                                <!-- 绑定到 overcharge_voltage_warning 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.overcharge_voltage_warning"
                                  :min="4.0"
                                  :max="5.0"
                                  :step="0.01"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首充截止电压_危险(V):</NText>
                                <!-- 绑定到 overcharge_voltage_danger 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.overcharge_voltage_danger"
                                  :min="4.0"
                                  :max="5.0"
                                  :step="0.01"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效_低值(%):</NText>
                                <!-- 绑定到 overcharge_efficiency_low 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.overcharge_efficiency_low"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首效_警告值(%):</NText>
                                <!-- 绑定到 overcharge_efficiency_warning 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.overcharge_efficiency_warning"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 容量衰减预警阈值 -->
                    <NCard title="容量衰减预警阈值">
                      <NSpace vertical size="medium">
                        <NAlert type="info" style="margin-bottom: 16px;">
                          <template #header>说明</template>
                          Cycle4是程序中固定的循环数，用于评估电池在第4个循环时的容量保持率和衰减情况。
                        </NAlert>
                        <NFormItem label="容量衰减阈值">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">Cycle4容量保持率_警告(%):</NText>
                                <!-- 绑定到 capacity_decay_cycle4_warning 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_decay_cycle4_warning"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">Cycle4容量保持率_危险(%):</NText>
                                <!-- 绑定到 capacity_decay_cycle4_danger 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_decay_cycle4_danger"
                                  :min="50"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">首放循环4差异(mAh/g):</NText>
                                <!-- 绑定到 capacity_decay_discharge_diff 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_decay_discharge_diff"
                                  :min="10"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 数据验证配置 -->
                    <NCard title="数据验证配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="循环次数限制">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最少循环次数要求:</NText>
                                <!-- 绑定到 min_cycles_required 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.min_cycles_required"
                                  :min="1"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">最大循环次数限制:</NText>
                                <!-- 绑定到 max_cycles_limit 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.max_cycles_limit"
                                  :min="100"
                                  :max="5000"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="数据范围验证">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">容量下限(mAh/g):</NText>
                                <!-- 绑定到 capacity_range_min 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_range_min"
                                  :min="0"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">容量上限(mAh/g):</NText>
                                <!-- 绑定到 capacity_range_max 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.capacity_range_max"
                                  :min="400"
                                  :max="600"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">电压下限(V):</NText>
                                <!-- 绑定到 voltage_range_min 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.voltage_range_min"
                                  :min="1.0"
                                  :max="2.5"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">电压上限(V):</NText>
                                <!-- 绑定到 voltage_range_max 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.voltage_range_max"
                                  :min="4.5"
                                  :max="6.0"
                                  :step="0.1"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">效率下限(%):</NText>
                                <!-- 绑定到 efficiency_range_min 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.efficiency_range_min"
                                  :min="0"
                                  :max="10"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">效率上限(%):</NText>
                                <!-- 绑定到 efficiency_range_max 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.efficiency_range_max"
                                  :min="100"
                                  :max="150"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">能量下限(mWh/g):</NText>
                                <!-- 绑定到 energy_range_min 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.energy_range_min"
                                  :min="0"
                                  :max="100"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">能量上限(mWh/g):</NText>
                                <!-- 绑定到 energy_range_max 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.energy_range_max"
                                  :min="1500"
                                  :max="2500"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>
                  </NSpace>
                </NTabPane>

                <!-- 绘图配置 -->
                <NTabPane name="plot" tab="绘图配置" display-directive="show:lazy">
                  <NSpace vertical size="large">
                    <!-- 绘图基本配置 -->
                    <NCard title="绘图基本配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="字体配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">字体族:</NText>
                                <!-- 绑定到 plot_font_family 参数 -->
                                <NSelect
                                  v-model:value="configStore.config.plot_font_family"
                                  class="config-input"
                                  :options="[
                                    { label: 'SimHei (黑体)', value: 'SimHei' },
                                    { label: 'Microsoft YaHei', value: 'Microsoft YaHei' },
                                    { label: 'Arial', value: 'Arial' },
                                    { label: 'Times New Roman', value: 'Times New Roman' }
                                  ]"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">字体大小:</NText>
                                <!-- 绑定到 plot_font_size 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.plot_font_size"
                                  :min="6"
                                  :max="20"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="图像质量">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">DPI:</NText>
                                <!-- 绑定到 plot_dpi 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.plot_dpi"
                                  :min="72"
                                  :max="600"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">图像宽度(英寸):</NText>
                                <!-- 绑定到 plot_figsize_width 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.plot_figsize_width"
                                  :min="6"
                                  :max="20"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">图像高度(英寸):</NText>
                                <!-- 绑定到 plot_figsize_height 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.plot_figsize_height"
                                  :min="4"
                                  :max="15"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="绘图后端">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">后端:</NText>
                                <!-- 绑定到 plot_backend 参数 -->
                                <NSelect
                                  v-model:value="configStore.config.plot_backend"
                                  class="config-input"
                                  :options="[
                                    { label: 'Agg (非交互式)', value: 'Agg' },
                                    { label: 'TkAgg', value: 'TkAgg' },
                                    { label: 'Qt5Agg', value: 'Qt5Agg' }
                                  ]"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 plot_interactive 参数 -->
                              <NCheckbox v-model:checked="configStore.config.plot_interactive">启用交互模式</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                      </NSpace>
                    </NCard>

                    <!-- 输出配置 -->
                    <NCard title="输出配置">
                      <NSpace vertical size="medium">
                        <NFormItem label="Excel输出配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">Excel写入引擎:</NText>
                                <!-- 绑定到 output_excel_engine 参数 -->
                                <NSelect
                                  v-model:value="configStore.config.output_excel_engine"
                                  class="config-input"
                                  :options="[
                                    { label: 'OpenPyXL (推荐)', value: 'openpyxl' },
                                    { label: 'XlsxWriter', value: 'xlsxwriter' }
                                  ]"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">数值保留小数位数:</NText>
                                <!-- 绑定到 decimal_places 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.decimal_places"
                                  :min="0"
                                  :max="10"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 include_charts 参数 -->
                              <NCheckbox v-model:checked="configStore.config.include_charts">包含图表</NCheckbox>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 compress_output 参数 -->
                              <NCheckbox v-model:checked="configStore.config.compress_output">压缩输出文件</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="图表配置">
                          <NGrid cols="1" y-gap="16">
                            <NGridItem>
                              <div class="config-row">
                                <NText class="config-label">图表DPI:</NText>
                                <!-- 绑定到 chart_dpi 参数 -->
                                <NInputNumber
                                  v-model:value="configStore.config.chart_dpi"
                                  :min="72"
                                  :max="600"
                                  class="config-input"
                                />
                              </div>
                            </NGridItem>
                            <NGridItem>
                              <!-- 绑定到 save_intermediate_results 参数 -->
                              <NCheckbox v-model:checked="configStore.config.save_intermediate_results">保存中间结果</NCheckbox>
                            </NGridItem>
                          </NGrid>
                        </NFormItem>
                        <NFormItem label="输出格式">
                          <div class="config-row">
                            <NText class="config-label">输出格式:</NText>
                            <!-- 绑定到 output_format 参数 -->
                            <NSelect
                              v-model:value="configStore.config.output_format"
                              class="config-input"
                              :options="[
                                { label: 'Excel (.xlsx) - 推荐', value: 'xlsx' },
                                { label: 'CSV (.csv)', value: 'csv' }
                              ]"
                            />
                          </div>
                        </NFormItem>
                      </NSpace>
                    </NCard>
                  </NSpace>
                </NTabPane>
              </NTabs>

              <!-- 底部操作按钮 -->
              <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid var(--border-color);">
                <NSpace>
                  <NButton type="primary">
                    <template #icon>
                      <NIcon><SettingsOutline /></NIcon>
                    </template>
                    保存配置
                  </NButton>
                  <NButton>
                    <template #icon>
                      <NIcon><PlayOutline /></NIcon>
                    </template>
                    重置为默认
                  </NButton>
                  <NButton>导出配置</NButton>
                  <NButton>导入配置</NButton>
                </NSpace>
              </div>
            </div>

            <!-- 数据处理面板 -->
            <div v-else-if="activePanel === 'process'" style="height: 100%; overflow-y: auto;">
              <!-- 页面标题 -->
              <div style="margin-bottom: 24px;">
                <NText style="font-size: 24px; font-weight: 600;">数据处理</NText>
                <NText depth="3" style="display: block; margin-top: 8px;">
                  选择数据文件夹，批量处理电池测试数据
                </NText>
              </div>

              <NGrid cols="1 800:2" x-gap="24" y-gap="24" responsive="screen">
                <!-- 左侧：文件选择和列表 -->
                <NGridItem>
                  <NSpace vertical size="large">
                    <!-- 文件夹选择 -->
                    <NCard title="数据文件夹" size="small">
                      <NSpace vertical>
                        <NSpace>
                          <NButton type="primary" @click="selectFolder">
                            <template #icon>
                              <NIcon><FolderOpenOutline /></NIcon>
                            </template>
                            选择文件夹
                          </NButton>
                          <NButton v-if="selectedFolder" @click="loadFileList">
                            <template #icon>
                              <NIcon><RefreshOutline /></NIcon>
                            </template>
                            刷新列表
                          </NButton>
                        </NSpace>
                        <NText v-if="selectedFolder" depth="3">
                          {{ selectedFolder }}
                        </NText>
                      </NSpace>
                    </NCard>

                    <!-- 文件列表 -->
                    <NCard title="文件列表" size="small">
                      <template #header-extra>
                        <NTag v-if="fileList.length > 0" type="info">
                          {{ fileList.length }} 个文件
                        </NTag>
                      </template>

                      <div style="height: 300px;">
                        <NScrollbar v-if="fileList.length > 0">
                          <NList>
                            <NListItem v-for="file in fileList" :key="file.path">
                              <NThing>
                                <template #avatar>
                                  <NIcon size="24" :color="getFileStatusColor(file.status)">
                                    <DocumentOutline v-if="file.status === 'pending'" />
                                    <TimeOutline v-else-if="file.status === 'processing'" />
                                    <CheckmarkCircleOutline v-else-if="file.status === 'completed'" />
                                    <CloseCircleOutline v-else-if="file.status === 'error'" />
                                  </NIcon>
                                </template>
                                <template #header>
                                  <NText>{{ file.name }}</NText>
                                </template>
                                <template #description>
                                  <NSpace>
                                    <NText depth="3">{{ formatFileSize(file.size) }}</NText>
                                    <NText depth="3">{{ file.lastModified.toLocaleDateString() }}</NText>
                                    <NTag :type="getFileStatusType(file.status)" size="small">
                                      {{ getFileStatusText(file.status) }}
                                    </NTag>
                                  </NSpace>
                                </template>
                              </NThing>
                            </NListItem>
                          </NList>
                        </NScrollbar>
                        <NEmpty v-else description="请先选择包含Excel文件的文件夹" />
                      </div>
                    </NCard>
                  </NSpace>
                </NGridItem>

                <!-- 右侧：处理控制和日志 -->
                <NGridItem>
                  <NSpace vertical size="large">
                    <!-- 处理控制 -->
                    <NCard title="处理控制" size="small">
                      <NSpace vertical>
                        <!-- 进度显示 -->
                        <div v-if="totalCount > 0">
                          <NText style="margin-bottom: 8px;">
                            处理进度: {{ processedCount }} / {{ totalCount }}
                          </NText>
                          <NProgress
                            type="line"
                            :percentage="processingProgress"
                            :status="getProgressStatus()"
                            :show-indicator="true"
                          />
                        </div>

                        <!-- 控制按钮 -->
                        <NSpace>
                          <NButton
                            type="primary"
                            :disabled="fileList.length === 0 || isProcessing"
                            @click="startProcessing"
                          >
                            <template #icon>
                              <NIcon><PlayOutline /></NIcon>
                            </template>
                            开始处理
                          </NButton>

                          <NButton
                            v-if="isProcessing && processingStatus === 'running'"
                            @click="pauseProcessing"
                          >
                            <template #icon>
                              <NIcon><PauseOutline /></NIcon>
                            </template>
                            暂停
                          </NButton>

                          <NButton
                            v-if="isProcessing"
                            type="error"
                            @click="stopProcessing"
                          >
                            <template #icon>
                              <NIcon><StopOutline /></NIcon>
                            </template>
                            停止
                          </NButton>
                        </NSpace>

                        <!-- 状态显示 -->
                        <NAlert v-if="processingStatus !== 'idle'" :type="getStatusAlertType()">
                          {{ getStatusText() }}
                        </NAlert>
                      </NSpace>
                    </NCard>

                    <!-- 处理日志 -->
                    <NCard title="处理日志" size="small">
                      <template #header-extra>
                        <NButton
                          size="small"
                          @click="clearLogs"
                          :disabled="processingLogs.length === 0"
                        >
                          <template #icon>
                            <NIcon><TrashOutline /></NIcon>
                          </template>
                          清空
                        </NButton>
                      </template>

                      <div style="height: 300px;">
                        <NScrollbar v-if="processingLogs.length > 0">
                          <NSpace vertical size="small">
                            <div
                              v-for="(log, index) in processingLogs"
                              :key="index"
                              style="padding: 8px; border-radius: 4px; background: var(--n-color-target);"
                            >
                              <NSpace align="center">
                                <NTag :type="getLogTagType(log.level)" size="small">
                                  {{ getLogLevelText(log.level) }}
                                </NTag>
                                <NTime :time="log.timestamp" format="HH:mm:ss" />
                                <NText>{{ log.message }}</NText>
                              </NSpace>
                            </div>
                          </NSpace>
                        </NScrollbar>
                        <NEmpty v-else description="暂无处理日志" />
                      </div>
                    </NCard>
                  </NSpace>
                </NGridItem>
              </NGrid>
            </div>

            <!-- 结果查看面板 -->
            <NCard v-else-if="activePanel === 'results'" title="结果查看">
              <NText>结果查看界面开发中...</NText>
            </NCard>

            <!-- 文档帮助面板 -->
            <NCard v-else-if="activePanel === 'docs'" title="使用文档">
              <NText>
                这里将显示详细的使用文档和帮助信息...
              </NText>
            </NCard>
          </div>
        </NLayoutContent>
      </NLayout>
    </NLayout>
  </NConfigProvider>
</template>

<style>
body {
  margin: 0;
  padding: 0;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Oxygen',
    'Ubuntu', 'Cantarell', 'Fira Sans', 'Droid Sans', 'Helvetica Neue',
    sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
}

#app {
  height: 100vh;
}

/* 全局样式优化 */
:deep(.n-card) {
  border-radius: 12px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
  transition: all 0.3s ease;
}

:deep(.n-card:hover) {
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
}

:deep(.n-form-item-label) {
  font-weight: 600;
  color: #333;
}

:deep(.n-input-number) {
  border-radius: 8px;
}

:deep(.n-select) {
  border-radius: 8px;
}

:deep(.n-input) {
  border-radius: 8px;
}

:deep(.n-checkbox) {
  margin: 4px 0;
}

:deep(.n-tabs-tab) {
  font-weight: 500;
  padding: 12px 20px;
}

:deep(.n-tabs-tab--active) {
  font-weight: 600;
}

/* 卡片标题样式 */
:deep(.n-card-header__main) {
  font-weight: 600;
  font-size: 16px;
  color: #2d3748;
}

/* 表单项间距 */
:deep(.n-form-item) {
  margin-bottom: 20px;
}

/* 网格间距优化 */
:deep(.n-grid) {
  gap: 16px;
}

/* 按钮样式优化 */
:deep(.n-button) {
  border-radius: 8px;
  font-weight: 500;
}

/* 标签页内容区域 */
:deep(.n-tab-pane) {
  padding: 20px 0;
}

/* 警告信息样式 */
:deep(.n-alert) {
  border-radius: 8px;
  margin-bottom: 16px;
}

/* 复选框组样式 */
:deep(.n-checkbox-group) {
  gap: 12px;
}

/* 选择器样式 */
:deep(.n-radio-group) {
  gap: 16px;
}

:deep(.n-radio) {
  margin-right: 16px;
}

/* 配置项对齐样式 */
.config-row {
  display: flex;
  align-items: center;
  gap: 16px; /* 固定间距，不使用space-between */
  width: 100%;
  min-height: 40px;
  padding: 4px 8px;
  border-radius: 6px;
  transition: all 0.3s ease;
}

.config-row:hover {
  background: linear-gradient(90deg, rgba(102, 126, 234, 0.05) 0%, transparent 100%);
  transform: translateX(2px);
  box-shadow: 0 2px 8px rgba(102, 126, 234, 0.1);
}

.config-label {
  min-width: 140px; /* 调整回合适的标签宽度 */
  text-align: left;
  font-weight: 500;
  color: #555;
  line-height: 34px; /* 与输入框高度对齐 */
  display: flex;
  align-items: center;
  flex-shrink: 0; /* 防止标签被压缩 */
}

.config-input {
  width: 140px;
  height: 34px; /* 固定高度确保一致性 */
  flex-shrink: 0; /* 防止输入框被压缩 */
}

/* 主容器样式 */
.main-container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 20px;
}

/* 侧边栏菜单样式优化 */
:deep(.n-menu-item) {
  border-radius: 8px;
  margin: 4px 8px;
}

:deep(.n-menu-item--selected) {
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
}

:deep(.n-menu-item--selected .n-menu-item-content-header) {
  color: white;
}

/* 布局优化 */
:deep(.n-layout-sider) {
  background: #f8fafc;
  border-right: 1px solid #e2e8f0;
}

:deep(.n-layout-content) {
  background: #ffffff;
}

/* 输入框聚焦效果 */
:deep(.n-input-number:focus-within) {
  box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.2);
}

:deep(.n-select:focus-within) {
  box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.2);
}

:deep(.n-input:focus-within) {
  box-shadow: 0 0 0 2px rgba(102, 126, 234, 0.2);
}
</style>