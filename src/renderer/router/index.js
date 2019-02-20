import Vue from 'vue'
import Router from 'vue-router'

/* Layout */
import Layout from '@/views/layout/Layout'

Vue.use(Router)

export default new Router({
  routes: [
    {
      path: '/landing',
      name: 'landing-page',
      component: require('@/components/LandingPage').default
    },
    {
      path: '',
      component: Layout,
      children: [
        {
          path: '/',
          component: () => import('@/views/excel-to-word'),
        },
        {
          path: '/home',
          component: () => import('@/views/Home'),
        },
      ]
    },
    { path: '*',  redirect: { path: '/404' } },
  ]
})
