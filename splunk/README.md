# Splunk recipe: member additions to watched AD groups

[Find-SwpADGroupMemberAdditions.spl](Find-SwpADGroupMemberAdditions.spl) searches 24 hours of Windows Security event 4728 (addition to a security-enabled **global** group). It does not cover removals or changes to local or universal groups. It reads events and an existing lookup; it does not write to Splunk or AD.

Replace the fictional index and confirm the sourcetype and field extractions using a known, authorized lab event. The recipe assumes XML fields `TargetSid` (group), `TargetUserName` (group name), `MemberSid` (added principal) and `SubjectUserName` (actor). No results does not prove no changes; check audit policy, event forwarding, field extraction, time range and lookup permissions first.

Provide an authorized CSV lookup with one row per exact group SID and a nonempty label. Fictional example:

```csv
group_sid,watch_label
S-1-5-21-10-20-30-1201,Example review group
```

If your field names or lookup alias differ, adjust the recipe before running it. Keep results and the lookup local to the authorized Splunk environment. Live Splunk validation pending.

- [Microsoft event 4728](https://learn.microsoft.com/en-us/windows/security/threat-protection/auditing/event-4728)
- [Splunk lookup command](https://help.splunk.com/en/splunk-enterprise/search/spl-search-reference/9.4/search-commands/lookup)
