import { createRouter, createWebHistory } from 'vue-router'
// @ 映射 src，路径和你目录完全匹配
import UploadReport from '@/views/uploadReport/uploadReport.vue'

const routes = [
    {
        path: '/uploadReport',
        name: 'UploadReport',
        component: UploadReport
    }
]

const router = createRouter({
    history: createWebHistory(),
    routes
})

export default router