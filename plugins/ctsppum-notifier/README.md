# ctsppum-notifier

Claude Code plugin for the CTSPPUM (UMMC Cardiothoracic Surgery) pre-admission messaging workflow.

| Skill | What it does |
|---|---|
| `ctsppum-patient-list` | Reads the CTSPPUM Google Calendar and builds `patient_list.xlsx` (Name, RN, Operation Planned, Date of Surgery, Contact Numbers, Old Date). Missing numbers come from iPesakit Patient Info. |
| `ctsppum-patient-notifier` | Opens the **CTSPPUM Patient Notifier** artifact, where the Excel becomes WhatsApp messages (admission or change-of-date, BM / EN / 中文). |

Notifier artifact: https://claude.ai/artifact/PN4YWSqYjDKnqUJcgBekuz (private until shared from its Share menu).
Source: `skills/ctsppum-patient-notifier/assets/ctsppum-notifier.html`.

## Install

```
/plugin marketplace add haosiang997/CTSPPUM
/plugin install ctsppum-notifier@ctsppum
```

Requires the Google Calendar connector for the patient list skill, and a browser session on
iPesakit for the phone-number fallback. Examples in the skills use fictional names and RNs.
