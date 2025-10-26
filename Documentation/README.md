# Gmail Attachment Saver - Process Documentation

This directory contains detailed process flow documentation for each major feature of the Gmail Attachment Saver application.

## 📋 Documentation Index

### Core Features Process Flows

| Feature | Document | Description |
|---------|----------|-------------|
| **Gmail Integration** | [gmail-integration-process.md](gmail-integration-process.md) | Contextual Gmail add-on integration, thread processing, and UI card construction |
| **Jira Integration** | [jira-integration-process.md](jira-integration-process.md) | Jira API connectivity, ticket retrieval, TPM workflows, and dynamic ticket details |
| **Smart Organization** | [smart-organization-process.md](smart-organization-process.md) | Automatic folder structure creation, hierarchical organization, and path generation |
| **Attachment Selection** | [attachment-selection-process.md](attachment-selection-process.md) | File type grouping, selection memory, UI generation, and state management |
| **Duplicate Handling** | [duplicate-handling-process.md](duplicate-handling-process.md) | Size-based duplicate detection, timestamped versioning, and conflict resolution |
| **Settings Management** | [settings-management-process.md](settings-management-process.md) | Secure credential storage, JQL configuration, validation, and connection testing |

## 🔄 Process Flow Overview

The Gmail Attachment Saver follows this high-level process flow:

```
1. Gmail Context Detection → 2. Settings Validation → 3. Jira Integration → 
4. Attachment Discovery → 5. User Selection → 6. Smart Organization → 
7. Duplicate Handling → 8. File Saving → 9. User Notification
```

### Detailed Flow Sequence

1. **📧 Gmail Integration**: User opens email → Add-on detects context → Loads attachments
2. **⚙️ Settings Validation**: Check Jira credentials → Validate connection → Load preferences  
3. **🎫 Jira Integration**: Fetch TPM tickets → Display dropdown → Handle ticket selection
4. **📎 Attachment Discovery**: Process email thread → Group by file type → Create selection UI
5. **☑️ User Selection**: User chooses attachments → Remember selections → Validate choices
6. **📁 Smart Organization**: Generate folder structure → Create hierarchy → Prepare target location
7. **🔍 Duplicate Handling**: Check existing files → Compare sizes → Handle conflicts intelligently
8. **💾 File Saving**: Save selected attachments → Apply versioning → Track results
9. **📢 User Notification**: Report success/failures → Show save statistics → Provide feedback

## 🎯 Key Integration Points

### Gmail ↔ Jira
- Thread ID links attachments to tickets
- Contextual ticket selection based on email content
- Persistent selection memory across Gmail navigation

### Jira ↔ Organization
- Ticket keys drive folder naming conventions
- Project hierarchy reflects Jira project structure
- TPM role filtering ensures relevant ticket access

### Selection ↔ Duplicate Handling
- Selection memory informs duplicate decisions
- User choices preserved across processing steps
- Smart conflict resolution based on user intent

## 📊 Error Handling Strategies

Each process includes comprehensive error handling:

- **Graceful Degradation**: System continues operating with reduced functionality
- **User Feedback**: Clear error messages with actionable guidance
- **Recovery Mechanisms**: Automatic retry logic and fallback options
- **Logging**: Detailed error tracking for troubleshooting

## 🚀 Performance Optimizations

Cross-cutting performance features:

- **Caching**: Settings, tickets, and folder references cached during sessions
- **Batch Processing**: Multiple files processed efficiently  
- **Lazy Loading**: UI components created on-demand
- **Resource Management**: Respects Google Apps Script execution limits

## 🔧 Technical Architecture

### Core Technologies
- **Google Apps Script**: Runtime environment and APIs
- **Gmail API**: Email and attachment access
- **Drive API**: File and folder management  
- **Jira REST API**: Ticket data and user information
- **CardService**: Gmail add-on UI framework

### Data Flow Patterns
- **Event-Driven**: User interactions trigger specific process flows
- **Stateless Processing**: Each operation independent and recoverable
- **Secure Storage**: User properties for settings and selection memory
- **API Integration**: RESTful communication with external services

## 📖 Reading Guide

### For Developers
1. Start with [gmail-integration-process.md](gmail-integration-process.md) for entry point understanding
2. Review [settings-management-process.md](settings-management-process.md) for configuration patterns
3. Study [jira-integration-process.md](jira-integration-process.md) for API integration examples

### For Technical Users
1. Begin with [smart-organization-process.md](smart-organization-process.md) for folder structure logic
2. Understand [attachment-selection-process.md](attachment-selection-process.md) for UI behavior
3. Learn [duplicate-handling-process.md](duplicate-handling-process.md) for file management

### For Troubleshooting
Each document includes dedicated error handling sections with:
- Common failure scenarios and causes
- Recovery strategies and workarounds  
- Logging and debugging information
- Performance considerations and limits

## 🔄 Process Dependencies

```
Settings Management ─┐
                     ├─→ Jira Integration ─┐
Gmail Integration ───┤                      ├─→ File Processing
                     └─→ Attachment Selection ─┘
                             │
                             ↓
                    Smart Organization ←─→ Duplicate Handling
```

## 📝 Documentation Standards

These process flows follow consistent formatting:
- **Overview**: High-level feature description
- **Process Flow**: Step-by-step technical implementation
- **Error Handling**: Failure scenarios and recovery
- **Performance**: Optimization strategies and limits
- **Integration**: Connection points with other features

---

*For the main project documentation, see the [root README.md](../README.md)*
