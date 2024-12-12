# Application Implementation Details

## Overview
This document outlines the implementation details of the hosted application, including authentication, logging, metrics, and secure access to secrets.

---

## Features Implemented

### 1. Application Hosting
- The application is successfully hosted in **Azure App Service**.

### 2. Single Sign-On (SSO)
- Configured **Microsoft Identity Provider** under **App Service Authentication** to enable secure access via Azure AD.

### 3. Version Control
- Source code is maintained in a **private GitHub repository**, ensuring secure and efficient version control.

### 4. Logging
- Logs are implemented and stored under the **root directory** of the application for monitoring and troubleshooting.

### 5. Metrics
- Application metrics are recorded in a **SharePoint list**, providing organized and accessible performance data.

### 6. Key Vault Integration
- **Azure Managed Identity** is enabled for secure interaction with Azure Key Vault.
- The **Key Vault Secrets User** role is assigned to the Managed Identity to control access to secrets.
- The Managed Identity has been added to the Key Vault's access policies.

---

## Additional Notes
- Ensure proper monitoring of logs and metrics to maintain application health.
- Regularly update and maintain access permissions for enhanced security.

---

## How to Contribute
For contributing, please submit a pull request to the **private GitHub repository**. Ensure adherence to the project's coding and documentation standards.

