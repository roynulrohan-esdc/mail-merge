<script lang="ts">
  export let label: string;
  export let checked: boolean = false;
  export let change: (event: Event) => void = undefined;
</script>

<div class="checkbox-field-inputs">
  <input type="checkbox" bind:checked on:change={change} {...$$restProps} />
  <p>{label}</p>
</div>

<style>
  .checkbox-field-inputs {
    display: flex;
    align-items: center;
    flex-direction: row !important;
    width: 100%;
  }

  .checkbox-field-inputs p {
    margin: 0;
    margin-left: 12px;
    font-size: 18px;
  }
</style>
