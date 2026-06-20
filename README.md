# SPA Document - Shell

This is an Angular application that generates and downloads a **SPA Deployment Request Form** as a `.docx` Word file.

## How it works

Fill in the following input fields on the form, then click **Download** to get the pre-populated Word document:

| Field | Description |
|---|---|
| Requested Date | Auto-filled with today's date |
| Artifact Name | Name of the deployment artifact |
| Artifact Version | Version of the artifact |
| Artifact Build Date | Date the artifact was built |
| Build Job Pipeline Id | GitLab pipeline ID for the build job |
| Git Branch | Source branch used for the build |
| Tag Number | Git tag associated with the release |
| Deployment Version | Version identifier for this deployment |
| Job Id | Job ID returned after deployment |
| Rollback Artifact Name | Artifact name to use if rollback is needed |
| Rollback Artifact Version | Version of the rollback artifact |

## Text styling in the document

The downloaded Word document preserves the following CSS styles applied in the preview:

- **Red text** (`.text-red`) — mandatory / important fields
- **Green text** (`.text-green`) — artifact name and post-deployment verification values
- **Blue text** (`.text-blue`) — hyperlinks and URLs
- **Yellow highlight** (`.background-yellow`) — highlighted instructions such as base URL notes

Styles are inherited through nested elements and rendered faithfully in the exported `.docx` file.

---

This project was generated using [Angular CLI](https://github.com/angular/angular-cli) version 20.0.0.

## Development server

To start a local development server, run:

```bash
ng serve
```

Once the server is running, open your browser and navigate to `http://localhost:4200/`. The application will automatically reload whenever you modify any of the source files.

