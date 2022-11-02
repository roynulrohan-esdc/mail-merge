<script>
  import { loadData, data } from "../stores/data";
  import { pageLoading } from "../stores/routes";

  $pageLoading = true;

  setTimeout(() => {
    if ($data.scenarioOne.length <= 0) loadData();

    $pageLoading = false;
  }, 200);
</script>

<div>
  <h1>Hello World</h1>

  <pre>
    <p>{JSON.stringify($data)}</p>
  </pre>
</div>
