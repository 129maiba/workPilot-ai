import axios from 'axios'
import { ElMessage } from 'element-plus'

const service = axios.create({
    baseURL: 'http://localhost:8080', // 后端接口地址
    timeout: 30000
})

// 请求拦截器（携带token）
service.interceptors.request.use(config => {
    // config.headers.Authorization = 'Bearer ' + localStorage.getItem('token')
    return config
})

// 响应拦截器
service.interceptors.response.use(res => {
    return res.data
}, err => {
    ElMessage.error(err.message)
    return Promise.reject(err)
})

export default service