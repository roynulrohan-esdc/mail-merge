import { get, readable, writable } from 'svelte/store';

export const settings = writable()

export const path = new ActiveXObject("WScript.Shell").CurrentDirectory //Current path of Mail Merge

export const lang = writable("en")

export const toggleLang = readable(() => {
    get(lang) === "en"
        ? lang.set("fr")
        : lang.set("en")
})