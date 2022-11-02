<script>
  import { loadData, data, failLoadData, failLoadMessage } from "../stores/data";
  import { changePage, pageLoading } from "../stores/routes";

  $pageLoading = true;

  setTimeout(() => {
    if (!$data) loadData();

    $pageLoading = false;
  }, 200);
</script>

<div>
  {#if $failLoadData}
    <h2>Failed to load data from excel</h2>

    <p><code>{$failLoadMessage}</code></p>

    <p>Please place <code>Scenario1.xlsx</code>, <code>Scenario2.xlsx</code> in the <code>/input/</code> directory and restart.</p>
  {:else}
    <h2>Menu</h2>

    <div>
      <table id="options-table" class="table mrgn-tp-lg">
        <thead>
          <tr>
            <th scope="col" class="col-sm-12">Option</th>
            <th scope="col" class="col-sm-1" />
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>Contact Employees</td>
            <td>
              <button
                class="btn btn-default"
                on:click={() => {
                  $changePage("review-data");
                }}
              >
                Select
              </button>
            </td>
          </tr>
          <tr>
            <td>Contact Managers</td>
            <td>
              <button
                class="btn btn-default"
                on:click={() => {
                  $changePage("review-data");
                }}
              >
                Select
              </button>
            </td>
          </tr>
          <tr>
            <td>Review information loaded by excel</td>
            <td>
              <button
                class="btn btn-default"
                on:click={() => {
                  $changePage("review-data");
                }}
              >
                Select
              </button>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
  {/if}
</div>

<style>
  h2 {
    color: grey;
    width: 100%;
    text-align: center;
  }
</style>
