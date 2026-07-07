<template>
  <div class="upload-excel">
    <el-upload
        ref="uploadRef"
        action=""
        :http-request="uploadExcel"
        :before-upload="beforeUpload"
        accept=".xlsx,.xls"
        :limit="1"
        drag
    >
      <el-icon class="el-icon--upload"><upload /></el-icon>
      <div class="el-upload__text">
        将Excel拖到此处，或<em>点击上传</em>
      </div>
      <template #tip>
        <div class="el-upload__tip">仅支持 xls / xlsx 文件，单文件最大10MB</div>
      </template>
    </el-upload>

    <!-- 加载弹窗 -->
    <el-dialog v-model="loadingVisible" title="导入中..." width="300px" :close-on-click-modal="false" :show-close="false">
      <div class="loading-box">
        <el-spinner size="40" />
        <p>正在解析Excel数据，请稍候...</p>
      </div>
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, defineEmits } from 'vue'
import { ElMessage } from 'element-plus'
import { Upload } from '@element-plus/icons-vue'
import axios from 'axios'

// 父组件接收事件
const emit = defineEmits(['success', 'error'])
const uploadRef = ref(null)
const loadingVisible = ref(false)

/**
 * 上传前校验文件
 */
const beforeUpload = (file) => {
  const isExcel = /\.(xlsx|xls)$/.test(file.name)
  const isLt10M = file.size / 1024 / 1024 < 10

  if (!isExcel) {
    ElMessage.error('只能上传 .xls / .xlsx 格式文件！')
    return false
  }
  if (!isLt10M) {
    ElMessage.error('文件不能超过10MB！')
    return false
  }
  return true
}

/**
 * 自定义上传请求
 */
const uploadExcel = async (params) => {
  loadingVisible.value = true
  const file = params.file
  // formData 传文件
  const formData = new FormData()
  formData.append('file', file)

  try {
    const res = await axios.post('/api/report/upload', formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      }
    })
    loadingVisible.value = false
    ElMessage.success('Excel导入成功')
    // 把后端返回的报表数据传给父组件
    emit('success', res.data)
    // 清空上传列表
    uploadRef.value.clearFiles()
  } catch (err) {
    loadingVisible.value = false
    const msg = err.response?.data?.message || '导入失败，请检查文件内容'
    ElMessage.error(msg)
    emit('error', err)
    uploadRef.value.clearFiles()
  }
}
</script>

<style scoped>
.loading-box {
  text-align: center;
  padding: 20px 0;
}
</style>