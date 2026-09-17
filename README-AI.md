# PROJECT_KNOWLEDGE.md

> Purpose:
> This document provides persistent business and architecture context for AI coding agents such as Codex.
>
> Before modifying architecture or core bill-conversion logic, read this document first.
>
> Confirmed decisions in this document should be treated as architectural constraints unless the user explicitly changes them.

---

# 1. Project Overview

The overall system currently contains or plans to contain several major domains:

- Office
  - Internal back-office system.
  - Used by internal employees.
  - Handles customers, contracts, bills, suppliers, franchisees, review, approval, publishing, etc.

- Portal
  - External-facing portal.
  - Used by customers and potentially suppliers/franchisees.
  - Contains customer-facing data and external workflow/process data.

- Bill Conversion
  - Converts supplier bills into standardized internal/customer bill templates.
  - Supplier/customer combinations may have different conversion rules.
  - Excel templates are important business rule carriers.

- ClientAI / ClientChatBot
  - Independent AI service intended for customer-facing AI assistance.
  - Office/Portal/other systems are data sources and capability providers.
  - ClientAI should not be tightly coupled to one existing application.

The system must support both Chinese and English data.

---

# 2. Core Bill Conversion Principle

The existing bill conversion system is the primary deterministic engine.

AI is NOT intended to replace the existing conversion engine.

Target architecture:

Original Supplier Bill
        |
        +------------------------+
        |                        |
        v                        v
Code Conversion Engine     AI Collection Engine
        |                        |
        v                        v
Code Collection Result     AI Collection Result
        |                        |
        +-----------+------------+
                    |
                    v
          Deterministic Validation
                    |
              PASS / FAIL
                    |
          Submit / Human Review

Important:

Code Engine = official production result.

AI Engine = independent second engine / auditor.

Validation Engine = deterministic comparison between Code and AI.

---

# 3. Existing Bill Conversion Flow

The existing system roughly follows:

Supplier Bill
    |
    v
Configuration
    |
    v
Mapping
    |
    v
Master Template
    |
    v
Conversion
    |
    v
Collection Sheet / PN Sheet
    |
    v
Validation
    |
    v
PDF / downstream workflow

Existing UI concepts include:

1. Configuration page
2. Mapping page
3. Conversion Result page
4. Final Collection/PN Sheet preview

Supplier × Customer combinations may have their own configuration.

Do NOT assume one universal supplier mapping works for all customers.

---

# 4. Master Template

The Master Template is an important part of the bill conversion architecture.

It defines:

- target Excel structure
- formulas
- collection fields
- presentation
- business calculation structure

Templates must be versioned.

Example:

Master Template v13
Master Template v14

AI configuration and Collection Schema should be associated with a specific template version.

Do NOT assume an AI schema remains valid after the Master Template changes.

---

# 5. Final Collection Sheet / PN Sheet

The most important output for AI validation is the final information collection sheet.

Most business information eventually flows into this sheet.

AI validation should primarily focus on:

"Is the information collected in the final Collection/PN Sheet correct?"

AI does NOT need to independently validate every sheet in the workbook.

Some special values are already controlled by:

- Mapping configuration
- manual configuration/popups
- external APIs
- Timeline or other business systems

These values do not necessarily need AI extraction from the supplier bill.

---

# 6. Existing Excel Calculation Architecture

Existing technologies include:

- LuckySheet
  - Front-end spreadsheet display/editing.

- HyperFormula
  - Front-end formula calculation.

- Spire
  - Backend Excel recalculation.
  - Excel rendering.
  - PDF generation.

- Microsoft Excel
  - Used as an authoritative compatibility/calculation reference during template validation.

- Apache POI
  - Used for formula scanning / compatibility analysis.

General principle:

Excel = calculation rules + template/presentation.

Database = workflow, state, configuration, audit, versioning, tracking.

HyperFormula / Spire = execution engines.

PDF = final rendered output.

Do NOT turn Excel into the workflow database.

---

# 7. Formula Governance

Formula compatibility is considered a critical risk.

Template validation should include formula scanning.

Formula functions may be classified into:

L1
- Known safe/common functions.
- Expected to work consistently across engines.

L2
- Functions requiring sample-based compatibility testing.

L3
- Unsupported/high-risk functions.
- Should normally be rejected.

Also reject or strictly control:

- macros
- external workbook links
- unsupported formulas
- overly complex template behavior

POI formula scanning answers:

"What formulas/functions are used?"

It does NOT prove the numerical result is correct.

---

# 8. Multi-Engine Excel Validation

The system uses or plans to use multiple calculation engines:

Excel
HyperFormula
Spire

Template validation should compare the same workbook using multiple engines.

Typical validation flow:

Master Template
      |
      +--> Excel calculation
      |
      +--> HyperFormula calculation
      |
      +--> Spire calculation
      |
      v
Compare important formula cells / PN Sheet output

The goal is to detect calculation engine inconsistencies before a template enters production.

The final Collection/PN Sheet is the highest-priority comparison target.

---

# 9. Pre-Submission Recalculation

Before an important bill is submitted/published/sent, the system should be capable of performing another deterministic calculation validation.

If critical calculation results differ between engines:

DO NOT automatically send the bill.

Instead:

- block submission
- record the mismatch
- notify or route to an administrator/reviewer

Calculation inconsistencies are considered critical failures.

---

# 10. AI Independent Collection Engine

The AI engine must be independent from the Code Conversion Engine.

This is a hard architectural requirement.

AI should receive:

1. Original supplier bill
2. Collection Schema / Master collection definition
3. Necessary extraction instructions

AI should NOT receive the Code Collection Result before performing extraction.

Wrong:

Supplier Bill
    |
Code Result
    |
AI checks Code Result

This is not independent validation.

Correct:

Supplier Bill
    |
    +--> Code Engine
    |
    +--> AI Engine
    |
Compare results

The purpose is to create two independent interpretations of the same source document.

---

# 11. AI Input

Supplier bills may include:

- digitally generated PDFs
- Excel-exported PDFs
- scanned PDFs
- complex tables
- Chinese bills
- English bills
- bilingual bills

For visual models, PDF pages may be rendered as images before sending them to the model.

The implementation should keep document preprocessing modular.

Do NOT tightly couple bill conversion logic to one model provider.

---

# 12. Collection Schema

AI must NOT freely decide the target structure.

The system provides a Collection Schema.

Example concept:

{
  "field": "housing_allowance",
  "label": "Housing Allowance",
  "type": "decimal",
  "required": false,
  "aliases": [
    "Housing Subsidy",
    "Accommodation Benefit",
    "住房补贴"
  ]
}

The schema may include:

- canonical field name
- display label
- description
- data type
- required flag
- aliases
- Chinese aliases
- English aliases
- business meaning
- validation rules
- importance
- source hints

The AI's responsibility is semantic extraction and mapping.

The schema defines the contract.

---

# 13. AI Output Contract

AI output must be structured data.

Prefer JSON / Structured Output.

Do NOT rely on natural-language prose as the primary machine-readable result.

Example:

{
  "records": [
    {
      "employee_id": "E001",
      "fields": [
        {
          "field": "housing_allowance",
          "value": 2000,
          "source": {
            "page": 2,
            "location": "table row 5, column Housing Allowance"
          },
          "confidence": 0.96
        }
      ]
    }
  ]
}

Important AI output concepts:

- employee/entity identifier
- canonical field
- extracted value
- source page
- source location
- confidence
- raw source text when useful

Provenance is important.

When Code and AI disagree, reviewers should be able to understand where the AI found the value.

---

# 14. AI Collection Sheet

AI should fundamentally produce canonical structured JSON.

The system may then:

AI JSON
   |
   v
Populate Master Template
   |
   v
AI Collection Sheet

The AI itself should NOT freely generate arbitrary Excel layouts.

The system controls the Excel template.

This keeps AI output deterministic and comparable.

---

# 15. Code vs AI Validation

Comparison should be FIELD-BASED.

Do NOT depend purely on Excel cell positions.

Preferred comparison key concept:

Employee / Entity
+
Canonical Field
+
Value

Example:

Employee E001
housing_allowance

Code = 2000
AI   = 2000

PASS

Example:

Employee E002
housing_allowance

Code = 2000
AI   = 2200

FAIL / REVIEW

---

# 16. Validation Dimensions

Validation should include at least:

1. Completeness
   - missing employee
   - missing field
   - missing line item

2. Value consistency
   - amount mismatch
   - text mismatch
   - date mismatch

3. Entity association
   - value assigned to wrong employee
   - wrong customer/entity relationship

4. Aggregate consistency
   - totals
   - employee count
   - line item totals

5. Unknown item detection
   - supplier bill contains an item not mapped into the Collection Schema

Do not blindly auto-map unknown financial items.

Unknown/high-impact items should normally require human confirmation.

---

# 17. Calculation Ownership

AI should NOT become the authoritative financial calculation engine.

Deterministic Code remains responsible for:

- VAT
- arithmetic
- totals
- rounding
- formulas
- deterministic business rules

AI is primarily responsible for:

- semantic understanding
- document interpretation
- field extraction
- identifying possible missing information
- detecting unexpected items

General principle:

AI understands.

Code calculates.

Validation compares.

Human resolves exceptions.

---

# 18. New / Unknown Line Items

Example:

Existing schema knows:

Base Salary
Housing Allowance
Transport Allowance

Supplier suddenly adds:

Remote Work Allowance

AI may detect:

unknown_line_item = "Remote Work Allowance"

The system should NOT automatically change official Mapping.

Preferred flow:

AI detects unknown item
        |
        v
Human review
        |
        +--> Ignore
        |
        +--> Map to existing canonical field
        |
        +--> Create new field/mapping
        |
        v
Update configuration/schema if approved

AI may suggest.

Human confirms.

Configuration becomes the official rule.

---

# 19. AI Configuration

Confirmed on 2026-09-16:

Use ONE shared AI collection engine for all Supplier × Customer combinations.

Different combinations are supported by configuration, NOT by creating a separate AI engine for each combination.

Configuration has two levels:

## 19.1 Global AI Engine Configuration

Manage shared runtime settings centrally:

- Provider
- Default model and model parameters
- Timeout and retry settings
- Structured Output mechanism
- Shared, versioned extraction instructions

Exact setting values and retry/fallback policies remain open until implementation and testing.

## 19.2 Supplier × Customer AI Validation Configuration

Manage business differences under the existing Supplier × Customer configuration:

- Target Collection Schema and schema version
- Associated Master Template version
- Field meanings and Chinese/English aliases
- Required and important fields
- Source hints and special extraction requirements
- Fields to validate or exclude, including special/API-controlled values
- Validation rules and validation level
- Applicable instruction version

Reuse common definitions where appropriate and configure only the differences.

Provider/model settings belong to the global configuration by default. Per-combination model overrides are NOT a confirmed requirement.

However, AI configuration should NOT simply be another ordinary Mapping table.

Suggested sections:

Basic Configuration
Master Template
Code Mapping
Special Data Source / API
AI Validation

Avoid giving users a completely unrestricted prompt box.

Prefer structured configuration that generates controlled prompts/instructions.

All combinations use the same execution flow:

Original Supplier Bill + Resolved Collection Schema / Extraction Instructions
    -> Shared AI Collection Engine
    -> Canonical Structured JSON
    -> Deterministic Code vs AI Validation

The AI input must still satisfy the independence requirements in sections 10 and 29. A code-generated intermediate Excel file must NOT be the sole source for independent AI extraction.

Prefer configuration for supplier-specific differences. If a difference cannot be expressed safely, assess a narrowly scoped extension rather than copying the entire engine.

A first pilot may cover only one Supplier × Customer × Template Version, but must use this shared engine architecture from the start.

---

# 20. Model Provider Abstraction

Do not hard-code the entire AI architecture around one provider/model.

Use an abstraction similar to:

AIProvider
    |
    +--> OpenAI
    +--> Future Provider
    +--> Local Model
    +--> Other Provider

Configuration should be able to represent:

provider
model
model parameters
schema version
instruction version

The business logic should depend on an AI interface, not a specific model SDK.

---

# 21. Current AI Model Strategy

Current model discussions are exploratory and are NOT permanent architecture decisions.

Models being considered include multimodal low-cost models such as:

- GPT-5.6 Luna
- GPT-4o mini
- GPT-4.1 mini
- other future/local multimodal models

Do NOT hard-code GPT-5.6 Luna into the domain model.

The model must remain configurable.

Model selection should ultimately be based on benchmark results.

---

# 22. AI Benchmark Strategy

Initial testing scale is approximately <= 100 real bills.

The test set should contain variety:

- Chinese
- English
- bilingual
- different suppliers
- different customers
- different layouts
- scanned bills
- digitally generated PDFs
- complex tables
- layout changes

Use the SAME test documents and SAME schema when comparing models.

Measure:

- field accuracy
- amount accuracy
- employee/entity mapping accuracy
- missing-field rate
- unknown-item detection
- valid JSON rate
- automatic PASS rate
- disagreement categories
- latency
- cost

Do NOT select a production model based only on model branding or benchmark marketing.

Use real bill data.

---

# 23. Cross-Language Semantic Mapping

Chinese/English semantic mapping is important.

Example:

Housing Allowance
Housing Subsidy
Accommodation Benefit
住房补贴

may all represent the same canonical business field.

The Collection Schema should support aliases and multilingual descriptions.

AI should map supplier terminology into canonical fields.

---

# 24. AI Validation UI

Conversion Result should be capable of displaying:

Code Engine        ✓
AI Engine          ✓
Validation         PASS / FAIL

When differences exist, show useful information such as:

Employee
Field
Code Value
AI Value
AI Source Page
AI Source Location
Confidence
Status

Users should be able to understand WHY validation failed.

Do not display only:

"AI validation failed."

---

# 25. Suggested Processing State

Example state flow:

CONVERSION_STARTED
      |
CODE_CONVERSION_COMPLETED
      |
CODE_VALIDATION_COMPLETED
      |
AI_COLLECTION_STARTED
      |
AI_COLLECTION_COMPLETED
      |
CODE_AI_VALIDATION
      |
      +--> PASS
      |
      +--> REVIEW_REQUIRED
      |
      +--> FAILED

Important:

AI failure should be distinguishable from business-data mismatch.

Example:

AI_MODEL_ERROR

is different from:

AI_CODE_VALUE_MISMATCH

---

# 26. Auditability

AI execution should be auditable.

Consider storing:

- conversion job id
- supplier/customer
- template version
- schema version
- AI provider
- AI model
- instruction/prompt version
- request timestamp
- latency
- token usage
- estimated cost
- AI result
- validation result
- reviewer action

Do not rely only on application logs.

Important AI decisions should be traceable.

---

# 27. Suggested Data Model Concepts

Potential entities/tables:

conversion_job

code_collection_result

ai_collection_result

validation_result

ai_model_run

ai_provider_config

collection_schema

collection_schema_version

template_version

mapping_version

Exact table names are implementation decisions.

Do NOT create duplicate tables if equivalent structures already exist.

Inspect the current repository/database first.

---

# 28. Existing System Must Be Preserved

Before implementing AI features:

READ THE EXISTING CODE.

Do NOT assume the architecture from this document exactly matches class/table names in the repository.

Map these concepts onto existing implementation.

Avoid unnecessary rewrites.

Especially:

DO NOT rewrite existing bill conversion logic just to introduce AI.

Prefer additive integration.

Confirmed scope on 2026-09-16:

Extracting shared logic from the existing supplier/region-specific Code Conversion Engines is a future maintenance objective, NOT part of the current AI integration scope.

Preserve existing special business behavior. Any later consolidation should be incremental and verified against representative bills and regression expectations.

---

# 29. Hard Rules for AI Implementation

These rules are considered confirmed unless explicitly changed:

1. AI does NOT replace Code Conversion Engine.

2. Code Result remains the official production result.

3. AI performs independent extraction.

4. AI must NOT see Code Collection Result before independent extraction.

5. AI output must be structured.

6. AI must use the Collection Schema.

7. AI should primarily validate the final Collection/PN Sheet.

8. Special/API-controlled values do not automatically require AI validation.

9. Code owns deterministic financial calculations.

10. Validation between Code and AI must be deterministic.

11. Unknown financial items must not silently modify Mapping.

12. AI must not directly modify official bill data.

13. Provider/model must remain configurable.

14. Template/schema versions must be traceable.

15. Existing conversion logic should not be rewritten without explicit approval.

16. Use one shared AI collection engine, configured for each Supplier × Customer combination.

17. Centralize AI runtime configuration; keep combination-specific business extraction and validation configuration separate.

18. Do not bundle existing Code Conversion Engine consolidation into the current AI integration work.

---

# 30. Office / Portal Data Architecture

Long-term data architecture principle:

Office = source of truth for official business data.

Portal = external projection + external workflow/process data.

Official data includes concepts such as:

- customers
- contracts
- approved/effective bills
- internal review/publishing state

Long-term architecture may become:

               Shared Identity
                    |
          +---------+---------+
          |                   |
        Office              Portal
          |                   |
     Office DB            Portal DB
          |
   Official Source
      of Truth

Portal should receive only the information required externally.

Do NOT blindly replicate the entire Office database into Portal.

---

# 31. Database Migration Strategy

Current/near-term strategy is intentionally phased.

PHASE 1:

Prioritize delivery of supplier/franchisee functionality.

- unified users/roles/menu where appropriate
- avoid premature master/slave business database split
- avoid unnecessary cross-database synchronization
- maintain clean module boundaries
- clearly distinguish states such as:
  - draft
  - submitted
  - published

Goal:

Finish and stabilize supplier/franchisee workflows first.

PHASE 2:

After workflows stabilize:

Office DB
- official full business data
- internal processing
- review
- publishing

Portal DB
- customer-facing projection
- supplier/franchisee draft/process data

Shared user database
- unified identity/account system

Official business truth remains in Office.

Synchronization is NOT full database replication.

---

# 32. Shared Identity / Authorization

A common user/identity database may be shared between platforms.

Authorization must still understand platform/context.

Examples:

Office User
Portal Customer
Supplier User
Franchisee User

Shared identity does NOT mean shared business permissions.

Authentication and authorization are separate concerns.

---

# 33. ClientAI / ClientChatBot

ClientAI should be an independent service.

Do NOT build it as a hard-coded feature that only understands Portal.

Conceptual modules:

gateway

orchestrator

customer-context

tools

tool registry

llm

connectors/portal

Future connectors may include:

Office
Outlook
nnroad
other internal systems

Existing systems are:

data sources
+
capability providers

ClientAI is the orchestration layer.

---

# 34. Client Isolation

ClientAI must maintain strict customer isolation.

A customer must only be able to access information they are authorized to access.

Never rely on the LLM itself to enforce tenant isolation.

Authorization must be deterministic and enforced by backend tools/connectors.

Customer context should be derived from authenticated identity.

Do NOT allow the model to freely specify arbitrary customer IDs to access data.

---

# 35. ClientAI Tool Architecture

The LLM should access business data through controlled Tools.

Example:

User:
"What is the status of my bill?"

        |
        v
ClientAI
        |
        v
BillStatusTool
        |
        v
Portal Connector
        |
        v
Authorized Portal API

The model should not directly query arbitrary production databases.

Tools provide the security boundary.

---

# 36. ClientAI MVP

An earlier desired MVP question was:

"Where is my order?"

However, analysis found that existing Portal/Office/nnroad systems did not have an Order/Tracking API at that stage.

A faster MVP alternative discussed was:

"Bill Status"

because Portal already had bill-related APIs/customer isolation.

This should be treated as an implementation planning decision that may need to be rechecked against the CURRENT repository.

Do NOT assume the old API situation is still true.

Inspect current code before implementing.

---

# 37. Security Principles

Important principles across the project:

- tenant/customer isolation
- least privilege
- auditable actions
- deterministic authorization
- no arbitrary production DB access by LLM
- controlled AI tools
- traceable model execution
- clear Office/Portal trust boundaries
- avoid unnecessary external data exposure

AI must operate inside application security boundaries.

AI must not become the security boundary itself.

---

# 38. Development Instructions for Codex

Before implementing a requested feature:

1. Read this document.

2. Inspect relevant existing code.

3. Identify existing:
   - modules
   - database entities
   - APIs
   - conversion pipeline
   - template/version system
   - mapping system
   - validation logic

4. Reuse existing architecture whenever reasonable.

5. Do not duplicate existing capabilities.

6. Do not silently change confirmed architecture.

7. If implementation requires changing a Hard Rule:
   STOP and explain why.

8. Clearly distinguish:
   - existing implementation
   - proposed change
   - assumptions
   - unresolved business questions

9. Prefer small, reversible changes.

10. Preserve backward compatibility where practical.

---

# 39. Before Writing Code

For large architectural changes, first produce:

A. Current State

B. Relevant Existing Modules

C. Proposed Change

D. Files / Modules Affected

E. Database Changes

F. API Changes

G. Risks

H. Migration / Compatibility

I. Test Plan

Then implement after the architecture is understood.

Do not begin by rewriting large modules.

---

# 40. AI Feature Testing Requirements

AI implementation tests should cover:

Normal cases

Missing fields

Unknown fields

Chinese field names

English field names

Bilingual documents

Different employee ordering

Different row ordering

Repeated line items

Decimal differences

Rounding differences

Multiple-page documents

Scanned documents

AI timeout

Invalid JSON

Provider failure

Low-confidence extraction

Template/schema version mismatch

Code vs AI mismatch

The system should fail safely.

---

# 41. Key Architectural Philosophy

The project follows this principle:

            AI
      Semantic Understanding
             |
             v
Supplier Bill --------> Structured Data
             |
             v
           Code
 Deterministic Calculation
             |
             v
        Validation
             |
             v
        Human Review
       when necessary

In short:

AI understands.

Code calculates.

Validation proves consistency.

Human resolves ambiguity.

---

# 42. Open Questions

The following should NOT automatically be treated as finalized requirements:

- final production AI provider/model
- final confidence threshold
- exact automatic PASS threshold
- exact AI retry/fallback strategy
- exact database table names
- whether every PDF is converted to images
- whether native PDF input will be used
- exact treatment of low-confidence values
- whether per-combination model overrides are needed beyond the global default
- exact handling of AI service outages
- final ClientAI MVP if APIs have changed

These should be decided based on implementation constraints and real testing.

---

# 43. Final Instruction to Coding Agent

Do not redesign the system simply because another architecture appears cleaner.

This project already has production/business constraints.

Your job is to:

UNDERSTAND existing implementation
        +
PRESERVE confirmed architecture
        +
ADD the requested capability
        +
MINIMIZE regression risk

When this document conflicts with current code:

Do not silently choose one.

Identify the conflict and report it before making a major architectural change.
