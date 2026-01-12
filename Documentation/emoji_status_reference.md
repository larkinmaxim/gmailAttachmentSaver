# Emoji Status Reference Guide

This document provides a complete reference of all emojis used in the **Select Active Ticket** dropdown to indicate ticket/project status.

## Status Emoji Table

| Emoji | Status Name | Description |
|:-----:|-------------|-------------|
| 🔧 | HYPERCARE | Project is in hypercare phase |
| 🔍 | HYPERCARE (WITH CHECK) | Project is in hypercare with additional checks required |
| 📋 | Order received | Order has been received and logged |
| 🧪 | Test system available | Test environment is ready for use |
| 🧪 | Test System Available (Implementation Order) | Test environment ready for implementation order |
| 🚀 | Project go-live/productive start | Project has gone live or started production |
| 👤 | System Design Assigned | User has been assigned to system design |
| 📨 | System Design Order Received | System design order has been received |
| 🎨 | System Design Started | System design work has begun |
| 👨‍💻 | Implementation Assigned | User has been assigned to implementation |
| 🧑‍💼 | Implementation Order Assigned | User assigned to manage the implementation order |
| 🔨 | Implementation Started | Implementation work has begun |
| ✅ | Requirements Clarified | All requirements have been clarified |
| 🔍 | Handover Check Needed | Handover check is required |
| 🟢 | LIVE SYSTEM AVAILABLE | Live system is available and operational |
| 📝 | *Default* | Any other status not listed above |

## Quick Visual Reference

### Project Lifecycle Stages
- 📋 → Order received
- 📨 → Design order received  
- 👤 → Design assigned
- 🎨 → Design started
- 👨‍💻 / 🧑‍💼 → Implementation assigned
- 🔨 → Implementation started
- 🧪 → Test system available
- ✅ → Requirements clarified
- 🔍 → Handover check needed
- 🚀 → Go-live/productive start
- 🔧 → Hypercare
- 🟢 → Live system available

## Source Code Reference

The emoji mappings are defined in `Code.js` in the `getStatusEmoji()` function (lines 4041-4061).

```javascript
function getStatusEmoji(status) {
  var statusEmojis = {
    'HYPERCARE': '🔧',
    'HYPERCARE (WITH CHECK)': '🔍',
    'Order received': '📋',
    'Test system available': '🧪',
    'Test System Available (Implementation Order)': '🧪',
    'Project go-live/productive start': '🚀',
    'System Design Assigned': '👤',
    'System Design Order Received': '📨',
    'System Design Started': '🎨',
    'Implementation Assigned': '👨‍💻',
    'Implementation Order Assigned': '🧑‍💼',
    'Implementation Started': '🔨',
    'Requirements Clarified': '✅',
    'Handover Check Needed': '🔍',
    'LIVE SYSTEM AVAILABLE': '🟢'
  };
  
  return statusEmojis[status] || '📝';
}
```

---
*Last updated: January 2026*
