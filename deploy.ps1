# Deploy current folder to Cloud Run (production). Env vars of the service are preserved.
gcloud run deploy consumption-web `
  --source . `
  --project consumption-web `
  --region europe-north1 `
  --allow-unauthenticated `
  --memory 2Gi --cpu 1 --timeout 600 --concurrency 1
