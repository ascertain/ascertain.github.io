---
layout: post
title: "Database Selection in Cloud: When to Use What"
date: 2026-05-18
author: Mohammad Kashif
tags: [kashif, mohammad-kashif, kashif-subrati, kashif-cloud-engineer, kashif-devops, cloud-database, bigquery, firestore, cloud-storage, sql, gcp, aws, azure, data-architecture, cloud-architecture, database-selection, data-engineering, cloud-native, serverless, data-lake, data-warehouse, nosql, relational-database, s3, blob-storage, dynamodb, cosmos-db, rds, cloud-sql, redshift, snowflake, aurora, devops, cloud-strategy]
categories: [cloud, architecture, kashif-blog]
description: "A concise guide to choosing the right database or storage service in the cloud — BigQuery vs Firestore vs SQL vs Cloud Storage. With cross-cloud mapping for GCP, AWS, and Azure. By Mohammad Kashif Subrati."
keywords:
  - Mohammad Kashif
  - Kashif Subrati
  - Mohammad Kashif Subrati
  - Kashif cloud engineer
  - Kashif DevOps
  - Kashif data engineering
  - cloud database selection guide
  - when to use BigQuery
  - when to use Firestore
  - when to use Cloud Storage
  - BigQuery vs Firestore vs SQL
  - GCP AWS Azure database comparison
  - cloud storage for audit data
  - data lifecycle management cloud
  - choosing right cloud database
  - cloud architecture best practices
  - serverless database options
  - transactional data cloud storage
---

# Database Selection in Cloud: When to Use What

Choosing the right storage or database service is one of the most impactful architectural decisions in cloud-native systems. Pick wrong and you'll pay — in cost, latency, or engineering headaches.

Here's a clear, decision-driven guide.

---

## The Decision Matrix

| **Use Case** | **GCP** | **AWS** | **Azure** |
|---|---|---|---|
| Analytics / Data Warehouse | BigQuery | Redshift / Athena | Synapse Analytics |
| Document / NoSQL (real-time) | Firestore | DynamoDB | Cosmos DB |
| Relational / Transactional | Cloud SQL / AlloyDB | RDS / Aurora | Azure SQL |
| Object / File Storage | Cloud Storage (GCS) | S3 | Blob Storage |

---

## When to Use a Data Warehouse (BigQuery / Redshift / Synapse)

**Use when:** You need to analyze large volumes of data — not serve it in real-time.

- Aggregations, reporting, dashboards
- Historical trend analysis
- Cross-dataset joins on millions of rows
- Batch analytics, BI tools (Looker, Power BI, QuickSight)

**Don't use for:** Real-time reads, single-row lookups, or transactional writes.

**Example:** You track every video call session (duration, participants, market, outcome). Store raw events, then query "What's the average call duration per market this quarter?" — that's a warehouse job.

---

## When to Use NoSQL / Document DB (Firestore / DynamoDB / Cosmos DB)

**Use when:** You need fast, flexible, real-time reads and writes — especially for apps.

- User profiles, session data, preferences
- Real-time sync (mobile/web apps)
- Key-value or document lookups
- Serverless, auto-scaling workloads

**Don't use for:** Complex joins, heavy aggregations, or strict relational integrity.

**Example:** Your app generates a meeting link on-the-fly. Store the link metadata (ID, created_by, expires_at, participants) in Firestore for instant retrieval when the user opens it.

---

## When to Use SQL / Relational DB (Cloud SQL / RDS / Azure SQL)

**Use when:** You need ACID transactions, relational integrity, and structured schemas.

- Financial transactions, order management
- Multi-table joins with foreign keys
- Systems where data consistency is non-negotiable
- Legacy app migrations

**Don't use for:** Unstructured data, massive analytics workloads, or file storage.

**Example:** An e-commerce checkout — inventory deduction, payment capture, and order creation must all succeed or all fail. That's a transaction. Use SQL.

---

## When to Use Object Storage (GCS / S3 / Blob Storage)

**Use when:** You need cheap, durable, scalable storage for files or archival data.

- PDFs, images, videos, logs
- Audit trails and compliance archives
- Data lake raw layer (land → process → warehouse)
- Backup and disaster recovery
- Long-term retention with lifecycle policies

**Don't use for:** Real-time queries, transactional operations, or fast key-value lookups.

---

## Real-World Example: Generate at Runtime, Store for Audit

**Scenario:** Your system generates a video meeting link in real-time, but you also need to store it for audit and compliance.

```
┌─────────────────────────────────────────────────────────┐
│  Request comes in → Generate link (runtime)             │
│                                                         │
│  1. Store link metadata in Firestore/DynamoDB           │
│     → fast retrieval, TTL for expiry                    │
│                                                         │
│  2. Push audit record to Cloud Storage (GCS/S3/Blob)    │
│     → cheap, durable, lifecycle-managed                 │
│                                                         │
│  3. Lifecycle policy:                                   │
│     - 30 days → Nearline/Infrequent Access              │
│     - 90 days → Coldline/Glacier                        │
│     - 365 days → Archive or delete                      │
│                                                         │
│  4. (Optional) Load into BigQuery for analytics         │
│     → "How many links generated per market per month?"  │
└─────────────────────────────────────────────────────────┘
```

**Why this works:**
- **Firestore/DynamoDB** handles the hot path — fast read/write for active links
- **Cloud Storage/S3** handles the cold path — audit records stored cheaply with automatic lifecycle transitions
- **BigQuery/Redshift** handles the analytics path — when you need insights across millions of records

---

## Storage Lifecycle: The Cost Saver

All three clouds offer lifecycle policies on object storage:

| **Tier** | **GCP** | **AWS** | **Azure** |
|---|---|---|---|
| Hot (frequent access) | Standard | S3 Standard | Hot |
| Warm (30+ days) | Nearline | S3 Infrequent Access | Cool |
| Cold (90+ days) | Coldline | S3 Glacier | Cold |
| Archive (365+ days) | Archive | S3 Glacier Deep Archive | Archive |

Set it once → data moves automatically → costs drop by 60-80% over time.

---

## Quick Decision Flowchart

```
Is it a file (PDF, image, log, archive)?
  → Cloud Storage / S3 / Blob

Do you need real-time reads/writes for an app?
  → Firestore / DynamoDB / Cosmos DB

Do you need ACID transactions with relational integrity?
  → Cloud SQL / RDS / Azure SQL

Do you need to analyze millions of rows for reporting?
  → BigQuery / Redshift / Synapse
```

---

## Key Takeaway

Don't force one service to do everything. Use each for what it's built for:

- **Warehouse** = analyze
- **NoSQL** = serve fast
- **SQL** = transact safely
- **Object Storage** = store cheaply, retain long

The best architectures combine them — generate at runtime (NoSQL), archive for audit (Object Storage), analyze for insights (Warehouse). Each layer does its job. Simple.

---

*Written by Mohammad Kashif — Technical Lead & Automation Specialist working with GCP, AWS, and cloud-native architectures at scale.*
