# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <info@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

import sqlite3
from typing import Any, Dict, Iterable, List, Optional, Set, Tuple, TypeAlias

ArticleRow: TypeAlias = Tuple[
    Optional[str],  # accession
    Optional[str],  # pmid
    Optional[str],  # title
    Optional[str],  # abstract
    Optional[str],  # journal
    Optional[str],  # year
    Optional[str],  # source
    Optional[str],  # date
    Optional[str],  # doi
    Optional[str],  # link
    Optional[str],  # volume
    Optional[str],  # issue
    Optional[str],  # page
    Optional[str],  # language
    Optional[str],  # publicationtype
    Optional[str],  # fulljournal
    Optional[str],  # issn
    Optional[str],  # si
    Optional[str],  # gr
    Optional[str],  # cin
    Optional[str],  # retrievedate
    Optional[str],  # provider
]
AuthorRow: TypeAlias = Tuple[str, int, str, Optional[str], Optional[str]]
MetadataDict: TypeAlias = Dict[str, Dict[str, Any]]


class MetadataStore:
    def __init__(self, db_path: str):
        self.conn = sqlite3.connect(db_path)
        self._apply_pragmas()
        self._ensure_schema()

    def _apply_pragmas(self):
        self.conn.execute("PRAGMA journal_mode=WAL;")
        self.conn.execute("PRAGMA synchronous=NORMAL;")
        self.conn.execute("PRAGMA temp_store=MEMORY;")

    def _ensure_schema(self):
        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS articles (
          accession       TEXT PRIMARY KEY,
          pmid            TEXT,
          title           TEXT,
          abstract        TEXT,
          journal         TEXT,
          year            TEXT,
          source          TEXT,
          date            TEXT,
          doi             TEXT,
          link            TEXT,
          volume          TEXT,
          issue           TEXT,
          page            TEXT,
          language        TEXT,
          publicationtype TEXT,
          fulljournal     TEXT,
          issn            TEXT,
          si              TEXT,
          gr              TEXT,
          cin             TEXT,
          retrievedate    TEXT,
          provider        TEXT
        );
        """)
        existing_cols = {
            row[1] for row in self.conn.execute("PRAGMA table_info(articles);").fetchall()
        }
        for col in ("si", "gr", "cin"):
            if col not in existing_cols:
                self.conn.execute(f"ALTER TABLE articles ADD COLUMN {col} TEXT;")

        self.conn.execute("""
        CREATE TABLE IF NOT EXISTS article_authors (
          accession    TEXT NOT NULL,
          author_order INTEGER NOT NULL,
          author_type  TEXT,
          full_name    TEXT,
          short_name   TEXT,
          PRIMARY KEY (accession, author_order),
          FOREIGN KEY (accession) REFERENCES articles(accession) ON DELETE CASCADE
        );
        """)

        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_articles_pmid ON articles(pmid);")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_articles_year ON articles(year);")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_articles_doi ON articles(doi);")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_articles_journal ON articles(journal);")

        self.conn.commit()

    def missing(self, accessions: Iterable[str]) -> List[str]:
        uniq = list({str(a) for a in accessions})
        if not uniq:
            return []

        q = ",".join("?" for _ in uniq)
        have = {
            r[0]
            for r in self.conn.execute(
                f"SELECT accession FROM articles WHERE accession IN ({q})",
                uniq,
            )
        }
        return [a for a in uniq if a not in have]

    def existing(self, accessions: Iterable[str]) -> List[str]:
        uniq = list({str(a) for a in accessions})
        if not uniq:
            return []

        q = ",".join("?" for _ in uniq)

        rows = self.conn.execute(
            f"SELECT accession FROM articles WHERE accession IN ({q})",
            uniq,
        ).fetchall()

        return [r[0] for r in rows]

    def get_metadata(self, accessions: Iterable[str]) -> MetadataDict:
        uniq = list({str(a) for a in accessions})
        if not uniq:
            return {}

        q = ",".join("?" for _ in uniq)

        article_rows = self.conn.execute(
            f"""
            SELECT
                accession, pmid, title, abstract, journal, year, source, date,
                doi, link, volume, issue, page, language, publicationtype,
                fulljournal, issn, si, gr, cin, retrievedate, provider
            FROM articles
            WHERE accession IN ({q})
            """,
            uniq,
        ).fetchall()

        result = {}
        for row in article_rows:
            (
                accession, pmid, title, abstract, journal, year, source, date,
                doi, link, volume, issue, page, language, publicationtype,
                fulljournal, issn, si, gr, cin, retrievedate, provider,
            ) = row

            result[accession] = {
                "article": {
                    "accession": accession,
                    "pmid": pmid,
                    "title": title,
                    "abstract": abstract,
                    "journal": journal,
                    "year": year,
                    "source": source,
                    "date": date,
                    "doi": doi,
                    "link": link,
                    "volume": volume,
                    "issue": issue,
                    "page": page,
                    "language": language,
                    "publicationtype": publicationtype,
                    "fulljournal": fulljournal,
                    "issn": issn,
                    "si": si,
                    "gr": gr,
                    "cin": cin,
                    "retrievedate": retrievedate,
                    "provider": provider,
                },
                "authors": [],
            }

        author_rows = self.conn.execute(
            f"""
            SELECT accession, author_order, author_type, full_name, short_name
            FROM article_authors
            WHERE accession IN ({q})
            ORDER BY accession, author_order
            """,
            uniq,
        ).fetchall()

        for accession, author_order, author_type, full_name, short_name in author_rows:
            if accession not in result:
                continue
            result[accession]["authors"].append(
                {
                    "author_order": author_order,
                    "author_type": author_type,
                    "full_name": full_name,
                    "short_name": short_name,
                }
            )

        return result

    def upsert_metadata(
        self,
        articles_rows: List[ArticleRow],
        authors_rows: List[AuthorRow],
    ) -> None:
        if not articles_rows and not authors_rows:
            return

        accessions: Set[str] = {row[0] for row in articles_rows if row[0] is not None}

        with self.conn:
            if articles_rows:
                self.conn.executemany("""
                INSERT INTO articles (
                    accession, pmid, title, abstract, journal, year, source, date,
                    doi, link, volume, issue, page, language, publicationtype,
                    fulljournal, issn, si, gr, cin, retrievedate, provider
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(accession) DO UPDATE SET
                    pmid            = excluded.pmid,
                    title           = excluded.title,
                    abstract        = excluded.abstract,
                    journal         = excluded.journal,
                    year            = excluded.year,
                    source          = excluded.source,
                    date            = excluded.date,
                    doi             = excluded.doi,
                    link            = excluded.link,
                    volume          = excluded.volume,
                    issue           = excluded.issue,
                    page            = excluded.page,
                    language        = excluded.language,
                    publicationtype = excluded.publicationtype,
                    fulljournal     = excluded.fulljournal,
                    issn            = excluded.issn,
                    si              = excluded.si,
                    gr              = excluded.gr,
                    cin             = excluded.cin,
                    retrievedate    = excluded.retrievedate,
                    provider        = excluded.provider;
                """, articles_rows)

            if accessions:
                q = ",".join("?" for _ in accessions)
                self.conn.execute(
                    f"DELETE FROM article_authors WHERE accession IN ({q})",
                    list(accessions),
                )

            if authors_rows:
                self.conn.executemany("""
                INSERT INTO article_authors (
                    accession, author_order, author_type, full_name, short_name
                ) VALUES (?, ?, ?, ?, ?);
                """, authors_rows)

    def close(self) -> None:
        self.conn.close()


# Backward-compatible alias (legacy name).
metadataStore = MetadataStore


def enrich_metadata(data: MetadataDict) -> MetadataDict:
    for pmid, entry in data.items():
        article = entry.get("article", {})
        authors = entry.get("authors", [])

        title = article.get("title", "") or ""
        year = article.get("year", "") or ""
        source = (article.get("source", "") or "").strip()

        if authors:
            fa = authors[0]
            full_name = fa.get("full_name", "") or ""
            short_name = fa.get("short_name", "") or ""
            if full_name:
                firstauthorlastname = full_name.split(",", 1)[0]
            else:
                firstauthorlastname = short_name.split(" ", 1)[0] if short_name else ""
        else:
            full_name = ""
            firstauthorlastname = ""
            short_name = ""

        if len(authors) >= 2:
            firstauthorlastnameetal = f"{firstauthorlastname} et al."
        elif len(authors) == 1:
            firstauthorlastnameetal = firstauthorlastname
        else:
            cn = article.get("collaborator", "") or ""
            firstauthorlastnameetal = cn if cn else ""

        authoryear = f"{firstauthorlastnameetal}, {year}" if year else firstauthorlastnameetal
        cite = f"{authoryear}.\n{title}\n{source}".strip()
        cite_maincheckbox = f"{title}\n{source}".strip()

        article["firstauthor"] = short_name or firstauthorlastname
        article["firstauthorlastname"] = firstauthorlastname
        article["firstauthorlastnameetal"] = firstauthorlastnameetal
        article["authoryear"] = authoryear
        article["cite"] = cite
        article["cite_maincheckbox"] = cite_maincheckbox

    return data
