import { readable, writable } from 'svelte/store';
import Menu from '../pages/Menu.svelte'
import ReviewData from '../pages/ReviewData.svelte'

// SPA Application style routing for HTA
export const routes = {
    "menu": Menu,
    "review-data": ReviewData
}

export let currentRoute = "menu"

export const page = writable(Menu)

export const pageLoading = writable(false)

export const changePage = readable((route) => {
    if (route in routes) {
        page.set(routes[route])
        currentRoute = route
    }
})