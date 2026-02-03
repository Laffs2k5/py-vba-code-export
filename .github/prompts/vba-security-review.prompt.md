# VBA Macro Security Audit Guide

## Purpose

This document provides comprehensive instructions for AI agents conducting security audits of VBA (Visual Basic for Applications) code. The focus is on identifying and analyzing security risks, with particular emphasis on network communications, external dependencies, and potential attack vectors.

## Audit Scope

### Primary Security Concerns

1. **Network Communications**
   - Web server connections and HTTP/HTTPS requests
   - DNS lookups and domain resolution
   - FTP/SFTP transfers
   - API calls to external services
   - WebSocket connections

2. **External Tools & Processes**
   - Shell command execution (`CreateObject("WScript.Shell")`)
   - PowerShell invocations
   - External executable launching
   - System command execution
   - Process spawning and interaction

3. **External Libraries & Dependencies**
   - COM object references (especially non-standard)
   - ActiveX control usage
   - DLL loading and dynamic library imports
   - Third-party library integration
   - Unknown or suspicious object instantiation

4. **Data Handling & Leakage**
   - Credential storage and transmission
   - File system access and data exfiltration
   - Clipboard operations
   - Registry modifications
   - Environmental variable access

## Security Review Methodology

### Code Analysis Framework

#### Step 1: Identify Entry Points

- Examine macro triggers (Auto_Open, Workbook_Open, Document_Open)
- Review event handlers (button clicks, form submissions)
- Check for hidden or obfuscated code
- Identify all function and subroutine definitions
- Document all external dependencies declared at module level

#### Step 2: Map External Communication Patterns

```vba
// Patterns to search for:
- MSXML2.XMLHTTP, MSXML2.ServerXMLHTTP
- WinHTTP.WinHTTPRequest
- URLmon.URLDownloadToFile
- InternetOpenURL
- CreateObject with URL/HTTP patterns
- .Open(.GET/.POST methods
- .Send( methods
```

**Questions to ask:**

- What URLs/domains are being contacted?
- Is HTTPS enforced?
- Are certificates validated?
- Is authentication implemented?
- Are requests logged?

#### Step 3: Trace Process Execution

```vba
// Patterns to search for:
- WScript.Shell.Exec()
- WScript.Shell.Run()
- CallByName with suspicious targets
- Shell() function calls
- CreateObject("Shell.Application")
- GetObject with shell references
```

**Questions to ask:**

- What commands are executed?
- Are commands constructed dynamically?
- Is user input sanitized?
- Are execution parameters validated?
- Could command injection occur?

#### Step 4: Audit External References

```vba
// Declaration patterns to examine:
- Late binding with CreateObject()
- Early binding with GUID references
- References dialog dependencies
- ActiveX control inclusions
```

**Questions to ask:**

- Are all referenced libraries legitimate?
- Are versions pinned or flexible?
- Do references match known security advisories?
- Are deprecated/end-of-life libraries being used?
- Are COM objects from trusted vendors?

#### Step 5: Data Flow Analysis

- Trace user input through the code
- Identify where data is read from (files, web, clipboard, registry)
- Map where data is written to (files, web, network)
- Check for encryption in transit and at rest
- Identify sensitive data handling (credentials, PII, financial data)

### Risk Scoring

Classify findings by severity:

| Severity | Indicators | Examples |
| --- | --- | --- |
| **Critical** | Active security vulnerabilities, direct credential exposure, remote code execution paths | Executing user-provided shell commands, downloading and executing code, storing plaintext passwords |
| **High** | Network communication without validation, suspicious process execution, unknown COM objects | Unvalidated HTTPS connections, dynamic command construction, undefined external tools |
| **Medium** | Questionable patterns, inadequate input validation, risky API usage | Weak certificate validation, insufficient logging, use of deprecated functions |
| **Low** | Code quality concerns, best practice violations, informational findings | Hardcoded URLs, missing comments, inconsistent error handling |

## Security Review Checklist

### Network Communications

- [ ] Identify all network communication code
- [ ] Verify HTTPS usage for sensitive data
- [ ] Check certificate validation implementation
- [ ] Review authentication mechanisms
- [ ] Validate domain/URL whitelisting
- [ ] Examine request/response handling
- [ ] Check for hardcoded credentials
- [ ] Review timeout configurations
- [ ] Verify error handling for network failures
- [ ] Check for logging of network activity

### Process Execution

- [ ] Identify all shell/process execution code
- [ ] Verify command construction methods
- [ ] Check for dynamic command generation
- [ ] Validate input sanitization
- [ ] Review execution permissions
- [ ] Check for parameter validation
- [ ] Verify error handling
- [ ] Review privilege escalation risks
- [ ] Check for process isolation
- [ ] Document intended vs. potential uses

### External Dependencies

- [ ] Catalog all external references
- [ ] Verify vendor legitimacy
- [ ] Check for known vulnerabilities
- [ ] Review version pinning
- [ ] Check for deprecation warnings
- [ ] Verify security updates availability
- [ ] Document dependency purposes
- [ ] Check for unnecessary dependencies
- [ ] Review COM object instantiation
- [ ] Validate ActiveX controls

### Data Protection

- [ ] Identify sensitive data handling
- [ ] Check encryption implementation
- [ ] Verify secure data destruction
- [ ] Review access control mechanisms
- [ ] Check audit logging
- [ ] Validate file permissions
- [ ] Review clipboard access
- [ ] Check registry operations
- [ ] Verify environment variable usage
- [ ] Document data flow

## Common VBA Security Vulnerabilities

### 1. Unvalidated Network Requests

```vba
// VULNERABLE
Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
xmlHTTP.Open "GET", userProvidedURL, False
xmlHTTP.Send

// SECURE
Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
If ValidateURL(userProvidedURL) Then
    xmlHTTP.Open "GET", userProvidedURL, False
    xmlHTTP.setRequestHeader "User-Agent", "MyApp/1.0"
    xmlHTTP.Send
End If
```

### 2. Command Injection

```vba
// VULNERABLE
Set shell = CreateObject("WScript.Shell")
shell.Run "cmd.exe /c " & userInput

// SECURE - Avoid dynamic command construction
// Use parameterized approaches or whitelisted values
```

### 3. Suspicious Object Creation

```vba
// CONCERNING - Verify necessity
Set obj = CreateObject("MSXML2.XMLHTTP")
Set shell = CreateObject("WScript.Shell")
Set process = CreateObject("Shell.Application")

// Question: Why is this needed?
```

### 4. Hardcoded Credentials

```vba
// VULNERABLE
xmlHTTP.setRequestHeader "Authorization", "Basic dXNlcjpwYXNzd29yZA=="

// Use secure credential storage or environment variables
```

### 5. Uncontrolled File Operations

```vba
// VULNERABLE
URLDownloadToFile 0, urlString, appPath & "\downloaded.exe"
Shell appPath & "\downloaded.exe"

// RISK: Download and execute arbitrary code
```

## Required Documentation

For each identified security finding, provide:

1. **Location**: File name, procedure, line number (if available)
2. **Vulnerability Type**: Classification from this guide
3. **Severity**: Critical/High/Medium/Low
4. **Description**: What the code does and why it's concerning
5. **Risk Assessment**: Potential attack scenarios
6. **Remediation**: Recommended fix or mitigation strategy
7. **Evidence**: Relevant code snippet

## Reporting Standards

### Findings Should Include

- Clear, non-technical explanation of the risk
- Technical details for remediation
- Reference to specific code locations
- Severity justification
- Recommended actions with priority
- Risk score for prioritization

### Output Format

```markdown
### Finding: [Title]
**Severity**: [Critical/High/Medium/Low]
**Location**: [File/Macro/Line]
**Type**: [Network/Process/Dependency/Data]

**Description**:
[Clear explanation of the issue]

**Risk**:
[Potential impact and attack scenarios]

**Evidence**:
[Code snippet]

**Recommendation**:
[Specific fix or mitigation]
```

## Best Practices for Security-Focused Code Review

1. **Assume Malicious Intent**: Analyze code as if an attacker has control of inputs
2. **Follow the Data**: Trace sensitive information through the entire application flow
3. **Defense in Depth**: Look for multiple layers of validation and error handling
4. **Principle of Least Privilege**: Verify permissions are minimal necessary
5. **Explicit Over Implicit**: Question default behaviors and hidden operations
6. **Fail Securely**: Check error handling doesn't expose sensitive information
7. **Documentation**: Require clear comments on security-sensitive code
8. **Whitelisting**: Prefer explicit allowlists over blacklists
9. **Immutability**: Check for mutable global state and shared resources
10. **Separation of Concerns**: Isolate security-critical code from general logic

## Tools & Indicators to Monitor

### Suspicious Functions

- `Shell()`, `CreateObject()`, `GetObject()`
- `URLDownloadToFile()`, `InternetOpenURL()`
- `CallByName()`, `Application.Run()`, `Eval()`
- Registry manipulation functions
- File I/O operations with suspicious paths

### Suspicious Patterns

- Obfuscated or encoded strings
- Base64-encoded commands
- Dynamic code generation
- Hidden modules or sheets
- Unusual control flow
- Excessive use of reflection/late binding

## Escalation Criteria

Findings that warrant immediate escalation:

- Active malware signatures
- Confirmed data exfiltration code
- Remote code execution vulnerabilities
- Credential theft mechanisms
- Privilege escalation paths
- Evidence of obfuscation to hide malicious intent

## References & Resources

- [OWASP VBA Security Review Guide](https://owasp.org/)
- [Microsoft VBA Security Documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/)
- [CWE-94: Improper Control of Generation of Code](https://cwe.mitre.org/data/definitions/94.html)
- [CWE-78: Improper Neutralization of Special Elements used in an OS Command](https://cwe.mitre.org/data/definitions/78.html)

## Audit Completion

- [ ] All code sections reviewed
- [ ] External communication mapped
- [ ] Process execution verified
- [ ] Dependencies cataloged
- [ ] Data flow analyzed
- [ ] Findings documented
- [ ] Risk scores assigned
- [ ] Recommendations provided
- [ ] Report reviewed
- [ ] Executive summary completed
