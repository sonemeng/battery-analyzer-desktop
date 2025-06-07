/**
 * 配置管理存储 - 统一版本
 *
 * 与后端 modules/config_parser.py 的 Config 类完全匹配
 * 采用扁平化结构，确保前后端配置一致性
 */

import { defineStore } from 'pinia'
import { ref, computed, watch } from 'vue'

/**
 * 电池数据分析器配置接口
 * 与后端 Config 类完全匹配
 */
export interface BatteryAnalyzerConfig {
  // ===== 基础运行参数 =====
  input_folder: string
  output_folder: string
  outlier_method: 'boxplot' | 'zscore_mad'
  reference_channel_method: 'traditional' | 'pca' | 'retention_curve_mse'
  output_format: 'xlsx' | 'csv'

  // ===== Excel读取配置 =====
  excel_engine: 'calamine' | 'openpyxl'
  cycle_sheet_name: string
  test_sheet_name: string

  // ===== 1C阈值配置 =====
  ratio_threshold: number
  discharge_diff_threshold: number
  overcharge_threshold: number
  very_low_efficiency_threshold: number
  low_efficiency_threshold: number
  default_1c_cycle: number

  // ===== 异常检测配置 =====
  // 箱线图法
  boxplot_use_method: boolean
  boxplot_threshold_discharge: number
  boxplot_threshold_efficiency: number
  boxplot_shrink_factor: number

  // Z-score+MAD方法
  zscore_mad_constant: number
  zscore_min_mad_ratio: number
  zscore_threshold_discharge: number
  zscore_threshold_efficiency: number
  zscore_threshold_voltage: number
  zscore_threshold_energy: number
  zscore_use_time_series: boolean
  zscore_min_samples_for_stl: number
  zscore_generate_plots: boolean

  // ===== 运行时配置 =====
  max_iterations: number

  // ===== 参考通道选择配置 =====
  // PCA配置
  pca_default_features: string[]
  pca_n_components: number
  pca_visualization_enabled: boolean
  pca_safe_voltage_threshold: number

  // 容量保留率配置
  capacity_retention_enabled: boolean
  capacity_retention_min_cycles: number
  capacity_retention_max_cycles: number
  capacity_retention_cycle_step: number
  capacity_retention_interpolation_method: string
  capacity_retention_use_raw_capacity: boolean
  capacity_retention_use_weighted_mse: boolean
  capacity_retention_weight_method: string
  capacity_retention_weight_factor: number
  capacity_retention_late_cycles_emphasis: number
  capacity_retention_dynamic_range: boolean
  capacity_retention_min_channels: number
  capacity_retention_include_voltage: boolean
  capacity_retention_include_energy: boolean
  capacity_retention_capacity_weight: number
  capacity_retention_voltage_weight: number
  capacity_retention_energy_weight: number
  capacity_retention_voltage_column: string
  capacity_retention_energy_column: string

  // ===== 绘图配置 =====
  plot_font_family: string
  plot_font_size: number
  plot_backend: string
  plot_dpi: number
  plot_figsize_width: number
  plot_figsize_height: number
  plot_interactive: boolean

  // ===== 文件名解析配置 =====
  device_id_max_length: number
  default_channel: string
  batch_id_prefix: string
  device_id_prefix: string

  // ===== 程序运行配置 =====
  verbose: boolean
  chunk_size: number
  memory_limit_mb: number
  enable_progress_bar: boolean
  auto_open_results: boolean
  backup_original_data: boolean
  log_level: string

  // ===== 数据验证配置 =====
  min_cycles_required: number
  max_cycles_limit: number
  capacity_range_min: number
  capacity_range_max: number
  voltage_range_min: number
  voltage_range_max: number
  efficiency_range_min: number
  efficiency_range_max: number
  energy_range_min: number
  energy_range_max: number

  // ===== 输出配置 =====
  output_excel_engine: string
  include_charts: boolean
  chart_dpi: number
  save_intermediate_results: boolean
  compress_output: boolean
  decimal_places: number

  // ===== 异常数据阈值 =====
  abnormal_high_charge: number
  abnormal_low_charge: number
  abnormal_low_discharge: number

  // ===== 过充风险阈值 =====
  overcharge_voltage_warning: number
  overcharge_voltage_danger: number
  overcharge_efficiency_low: number
  overcharge_efficiency_warning: number

  // ===== 容量衰减预警阈值 =====
  capacity_decay_cycle4_warning: number
  capacity_decay_cycle4_danger: number
  capacity_decay_discharge_diff: number

  // ===== 测试模式配置 =====
  mode_patterns: string[]
  mode_one_c_modes: string[]
  mode_non_one_c_modes: string[]

  // ===== 文件系列标识配置 =====
  series_default: string
}

// 删除旧的嵌套配置接口，使用统一的扁平化配置

export const useConfigStore = defineStore('config', () => {
  // 默认配置 - 与后端 Config 类完全匹配
  const defaultConfig: BatteryAnalyzerConfig = {
    // ===== 基础运行参数 =====
    input_folder: '',
    output_folder: '',
    outlier_method: 'boxplot',
    reference_channel_method: 'retention_curve_mse',
    output_format: 'xlsx',

    // ===== Excel读取配置 =====
    excel_engine: 'calamine',
    cycle_sheet_name: 'Cycle',
    test_sheet_name: 'test',

    // ===== 1C阈值配置 =====
    ratio_threshold: 0.85,
    discharge_diff_threshold: 15,
    overcharge_threshold: 350,
    very_low_efficiency_threshold: 80,
    low_efficiency_threshold: 85,
    default_1c_cycle: 3,

    // ===== 异常检测配置 =====
    // 箱线图法
    boxplot_use_method: true,
    boxplot_threshold_discharge: 10,
    boxplot_threshold_efficiency: 3,
    boxplot_shrink_factor: 0.95,

    // Z-score+MAD方法
    zscore_mad_constant: 0.6745,
    zscore_min_mad_ratio: 0.01,
    zscore_threshold_discharge: 3.0,
    zscore_threshold_efficiency: 2.5,
    zscore_threshold_voltage: 3.0,
    zscore_threshold_energy: 3.0,
    zscore_use_time_series: true,
    zscore_min_samples_for_stl: 10,
    zscore_generate_plots: true,

    // ===== 运行时配置 =====
    max_iterations: 10,

    // ===== 参考通道选择配置 =====
    // PCA配置
    pca_default_features: ['首放', '首圈电压', 'Cycle4'],
    pca_n_components: 2,
    pca_visualization_enabled: true,
    pca_safe_voltage_threshold: 4.65,

    // 容量保留率配置
    capacity_retention_enabled: true,
    capacity_retention_min_cycles: 5,
    capacity_retention_max_cycles: 800,
    capacity_retention_cycle_step: 1,
    capacity_retention_interpolation_method: 'linear',
    capacity_retention_use_raw_capacity: true,
    capacity_retention_use_weighted_mse: true,
    capacity_retention_weight_method: 'linear',
    capacity_retention_weight_factor: 1.0,
    capacity_retention_late_cycles_emphasis: 2.0,
    capacity_retention_dynamic_range: true,
    capacity_retention_min_channels: 2,
    capacity_retention_include_voltage: true,
    capacity_retention_include_energy: true,
    capacity_retention_capacity_weight: 0.6,
    capacity_retention_voltage_weight: 0.1,
    capacity_retention_energy_weight: 0.3,
    capacity_retention_voltage_column: '当前电压保持',
    capacity_retention_energy_column: '当前能量保持',

    // ===== 绘图配置 =====
    plot_font_family: 'SimHei',
    plot_font_size: 10,
    plot_backend: 'Agg',
    plot_dpi: 300,
    plot_figsize_width: 10,
    plot_figsize_height: 6,
    plot_interactive: true,

    // ===== 文件名解析配置 =====
    device_id_max_length: 20,
    default_channel: 'CH-01',
    batch_id_prefix: 'BATCH-',
    device_id_prefix: 'DEVICE-',

    // ===== 程序运行配置 =====
    verbose: false,
    chunk_size: 50,
    memory_limit_mb: 500,
    enable_progress_bar: true,
    auto_open_results: false,
    backup_original_data: true,
    log_level: 'INFO',

    // ===== 数据验证配置 =====
    min_cycles_required: 2,
    max_cycles_limit: 1000,
    capacity_range_min: 0,
    capacity_range_max: 500,
    voltage_range_min: 2.0,
    voltage_range_max: 5.0,
    efficiency_range_min: 0,
    efficiency_range_max: 120,
    energy_range_min: 0,
    energy_range_max: 2000,

    // ===== 输出配置 =====
    output_excel_engine: 'openpyxl',
    include_charts: true,
    chart_dpi: 300,
    save_intermediate_results: false,
    compress_output: false,
    decimal_places: 2,

    // ===== 异常数据阈值 =====
    abnormal_high_charge: 380,
    abnormal_low_charge: 200,
    abnormal_low_discharge: 200,

    // ===== 过充风险阈值 =====
    overcharge_voltage_warning: 4.65,
    overcharge_voltage_danger: 4.7,
    overcharge_efficiency_low: 80,
    overcharge_efficiency_warning: 75,

    // ===== 容量衰减预警阈值 =====
    capacity_decay_cycle4_warning: 85,
    capacity_decay_cycle4_danger: 80,
    capacity_decay_discharge_diff: 50,

    // ===== 测试模式配置 =====
    mode_patterns: ['-0.1C-', '-0.5C-', '-1C-', '-BL-', '-0.33C-'],
    mode_one_c_modes: ['-1C-'],
    mode_non_one_c_modes: ['-0.1C-', '-BL-', '-0.33C-'],

    // ===== 文件系列标识配置 =====
    series_default: 'Q3'
    // 配置完成
  }

  // 从localStorage加载配置
  const loadConfigFromStorage = (): BatteryAnalyzerConfig => {
    try {
      const saved = localStorage.getItem('battery-analyzer-config')
      if (saved) {
        const parsed = JSON.parse(saved)
        console.log('✅ 从缓存加载配置:', Object.keys(parsed).length, '个参数')
        return { ...defaultConfig, ...parsed }
      }
    } catch (error) {
      console.warn('⚠️ 加载配置失败，使用默认配置:', error)
    }
    return JSON.parse(JSON.stringify(defaultConfig))
  }

  // 保存配置到localStorage
  const saveConfigToStorage = (configData: BatteryAnalyzerConfig) => {
    try {
      localStorage.setItem('battery-analyzer-config', JSON.stringify(configData))
      console.log('✅ 配置已保存到缓存')
    } catch (error) {
      console.error('❌ 保存配置失败:', error)
    }
  }

  // 当前配置状态（从缓存加载）
  const config = ref<BatteryAnalyzerConfig>(loadConfigFromStorage())

  // 计算属性
  const isModified = computed(() => {
    return JSON.stringify(config.value) !== JSON.stringify(defaultConfig)
  })

  // 监听配置变化并自动保存
  watch(config, (newConfig) => {
    saveConfigToStorage(newConfig)
  }, { deep: true })

  // 方法
  const resetToDefault = () => {
    config.value = JSON.parse(JSON.stringify(defaultConfig))
    saveConfigToStorage(config.value)
  }

  const updateConfig = (path: string, value: any) => {
    const keys = path.split('.')
    let current: any = config.value
    
    for (let i = 0; i < keys.length - 1; i++) {
      current = current[keys[i]]
    }
    
    current[keys[keys.length - 1]] = value
  }

  const getConfig = (path: string) => {
    const keys = path.split('.')
    let current: any = config.value
    
    for (const key of keys) {
      current = current[key]
    }
    
    return current
  }

  const exportConfig = () => {
    return JSON.stringify(config.value, null, 2)
  }

  const importConfig = (configJson: string) => {
    try {
      const importedConfig = JSON.parse(configJson)
      config.value = { ...defaultConfig, ...importedConfig }
      return true
    } catch (error) {
      console.error('导入配置失败:', error)
      return false
    }
  }

  // 转换为后端处理格式
  const toBackendConfig = () => {
    // 直接返回配置对象，因为现在结构完全匹配
    return { ...config.value }
  }

  // 从后端配置更新前端配置
  const fromBackendConfig = (backendConfig: Partial<BatteryAnalyzerConfig>) => {
    config.value = { ...config.value, ...backendConfig }
  }

  return {
    config,
    defaultConfig,
    isModified,
    resetToDefault,
    updateConfig,
    getConfig,
    exportConfig,
    importConfig,
    toBackendConfig,
    fromBackendConfig
  }
})
