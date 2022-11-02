import { readable, writable } from 'svelte/store';
import MailSelect from '../pages/MailSelect.svelte'

// SPA Application style routing for HTA
export const routes = {
    "mail-select": MailSelect
}

export let currentRoute = "mail-select"

export const page = writable(MailSelect)

export const pageLoading = writable(false)

export const changePage = readable((route) => {
    if (route in routes) {
        page.set(routes[route])
        currentRoute = route
    }
})