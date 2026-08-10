//! Production lifecycle for resumable worksheet-row projection.
//!
//! The cursor owns the worksheet entry stream and projector between pulls. The
//! surrounding [`XlsxZip`](crate::XlsxZip) continues to own the single logical
//! package operation so workbook dependencies, every row pull, and the sheet's
//! ancillary parts are charged to the same operation.

use std::collections::BTreeMap;
use std::io::BufReader;
use std::rc::Rc;

use ooxml_common::package_session::PackageEntryStream;

use crate::worksheet_projector::{
    ProjectedWorksheetRow, WorksheetProjectorItem, WorksheetRowProjector,
};
use crate::{Row, SharedString, XlsxZip};

/// Default semantic credit for one production pull. Rows are indivisible: the
/// cursor never splits a row to satisfy this credit.
pub(super) const WORKSHEET_CURSOR_PULL_ROWS: usize = 128;
pub(super) const WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES: usize = 1024 * 1024;

type ProductionProjector =
    WorksheetRowProjector<BufReader<PackageEntryStream>, Rc<[SharedString]>, Rc<[String]>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum WorksheetCursorState {
    Open,
    Finished,
    Failed,
    Canceled,
    Closed,
}

#[derive(Debug)]
pub(super) struct WorksheetCursorTail {
    pub(super) shell_xml: String,
    pub(super) row_heights: BTreeMap<u32, f64>,
}

#[derive(Debug)]
pub(super) enum WorksheetCursorPull {
    Rows {
        rows: Vec<Row>,
        /// Sum of internal standalone row-projection bytes. This is not a
        /// serialized wire size and therefore is not protocol byte credit.
        projected_bytes: usize,
    },
    Finished(WorksheetCursorTail),
}

/// An owned worksheet-entry cursor whose projector survives across pulls.
///
/// Row batches are provisional until [`WorksheetCursorPull::Finished`]. A
/// caller assembling retained state must not commit earlier batches: the ZIP
/// CRC and well-formed worksheet tail are validated only when the entry reaches
/// EOF. `cancel` and `close` are deliberately idempotent and immediately drop
/// the entry stream/decoder lease.
pub(super) struct WorksheetCursor {
    projector: Option<ProductionProjector>,
    pending_row: Option<ProjectedWorksheetRow>,
    pending_tail: Option<WorksheetCursorTail>,
    state: WorksheetCursorState,
}

impl WorksheetCursor {
    fn open_under_active_operation(
        archive: &mut XlsxZip,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<Self, String> {
        let entry = archive.active_operation()?.open_entry(part)?;
        let projector =
            WorksheetRowProjector::from_package_entry(entry, shared_strings, theme_colors)?;
        Ok(Self {
            projector: Some(projector),
            pending_row: None,
            pending_tail: None,
            state: WorksheetCursorState::Open,
        })
    }

    pub(super) fn pull(
        &mut self,
        max_rows: usize,
        target_projected_bytes: usize,
    ) -> Result<WorksheetCursorPull, String> {
        // `max_rows` is a hard semantic-unit limit and is clamped below. The
        // projected-byte argument is intentionally a soft internal batching
        // target: one indivisible row may cross it after passing the separate
        // 8 MiB hard row-projection cap. A future wire adapter must measure its
        // serialized payload independently and obey protocol `byteCredit`.
        if max_rows == 0 || target_projected_bytes == 0 {
            return Err("worksheet cursor pull limits must be greater than zero".to_string());
        }
        if self.state != WorksheetCursorState::Open {
            return Err(self.inactive_error());
        }
        if let Some(tail) = self.pending_tail.take() {
            self.state = WorksheetCursorState::Finished;
            self.projector.take();
            return Ok(WorksheetCursorPull::Finished(tail));
        }

        let row_limit = max_rows.min(WORKSHEET_CURSOR_PULL_ROWS);
        let mut rows = Vec::with_capacity(row_limit);
        let mut projected_bytes = 0usize;
        loop {
            let item = match self.pending_row.take() {
                Some(row) => Ok(WorksheetProjectorItem::Row(row)),
                None => self
                    .projector
                    .as_mut()
                    .expect("open worksheet cursor owns its projector")
                    .next_item(),
            };
            match item {
                Ok(WorksheetProjectorItem::Row(row)) => {
                    let next_bytes = projected_bytes.saturating_add(row.projected_bytes);
                    if !rows.is_empty() && next_bytes > target_projected_bytes {
                        self.pending_row = Some(row);
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                    projected_bytes = next_bytes;
                    rows.push(row.row);
                    if rows.len() == row_limit {
                        return Ok(WorksheetCursorPull::Rows {
                            rows,
                            projected_bytes,
                        });
                    }
                }
                Ok(WorksheetProjectorItem::Finished(tail)) => {
                    let tail = WorksheetCursorTail {
                        shell_xml: tail.shell_xml,
                        row_heights: tail.row_heights,
                    };
                    if rows.is_empty() {
                        self.state = WorksheetCursorState::Finished;
                        self.projector.take();
                        return Ok(WorksheetCursorPull::Finished(tail));
                    }
                    self.pending_tail = Some(tail);
                    return Ok(WorksheetCursorPull::Rows {
                        rows,
                        projected_bytes,
                    });
                }
                Err(error) => {
                    self.state = WorksheetCursorState::Failed;
                    self.projector.take();
                    self.pending_tail = None;
                    return Err(error.to_string());
                }
            }
        }
    }

    pub(super) fn cancel(&mut self) {
        if matches!(
            self.state,
            WorksheetCursorState::Canceled | WorksheetCursorState::Closed
        ) {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Canceled;
    }

    pub(super) fn close(&mut self) {
        if self.state == WorksheetCursorState::Closed {
            return;
        }
        self.projector.take();
        self.pending_row = None;
        self.pending_tail = None;
        self.state = WorksheetCursorState::Closed;
    }

    fn inactive_error(&self) -> String {
        let state = match self.state {
            WorksheetCursorState::Open => "open",
            WorksheetCursorState::Finished => "finished",
            WorksheetCursorState::Failed => "failed",
            WorksheetCursorState::Canceled => "canceled",
            WorksheetCursorState::Closed => "closed",
        };
        format!("worksheet cursor is {state}")
    }
}

impl XlsxZip {
    /// Open a persistent worksheet cursor only inside an explicitly-started
    /// package operation. This makes operation ownership visible at the factory
    /// boundary and prevents the lazy compatibility operation from escaping
    /// across production pulls.
    pub(super) fn open_worksheet_cursor(
        &mut self,
        part: &str,
        shared_strings: Rc<[SharedString]>,
        theme_colors: Rc<[String]>,
    ) -> Result<WorksheetCursor, String> {
        WorksheetCursor::open_under_active_operation(self, part, shared_strings, theme_colors)
    }
}

impl Drop for WorksheetCursor {
    fn drop(&mut self) {
        self.close();
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};

    use zip::write::SimpleFileOptions;

    use super::*;
    use crate::open_zip;

    const PART: &str = "xl/worksheets/sheet1.xml";

    fn worksheet(row_count: usize, tail: &str) -> String {
        let mut xml = String::from(
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
        );
        for row in 1..=row_count {
            xml.push_str(&format!(
                r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>row {row}</t></is></c></row>"#
            ));
        }
        xml.push_str(tail);
        xml
    }

    fn package(xml: &str) -> Vec<u8> {
        let mut bytes = Vec::new();
        {
            let mut writer = zip::ZipWriter::new(Cursor::new(&mut bytes));
            writer
                .start_file(
                    PART,
                    SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored),
                )
                .unwrap();
            writer.write_all(xml.as_bytes()).unwrap();
            writer.finish().unwrap();
        }
        bytes
    }

    fn corrupt_crc_consistently(bytes: &mut [u8]) {
        let wrong_crc = u32::from_le_bytes(bytes[14..18].try_into().unwrap()) ^ 0xffff_ffff;
        bytes[14..18].copy_from_slice(&wrong_crc.to_le_bytes());
        let central = bytes
            .windows(4)
            .position(|window| window == 0x0201_4b50u32.to_le_bytes())
            .expect("central directory header");
        bytes[central + 16..central + 20].copy_from_slice(&wrong_crc.to_le_bytes());
    }

    #[test]
    fn production_cursor_keeps_one_operation_across_multiple_atomic_pulls() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive
            .begin_operation("worksheet-cursor")
            .expect("operation starts");
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .expect("cursor opens");

        let mut batches = Vec::new();
        let mut inflated_snapshots = Vec::new();
        let (batches, tail) = loop {
            match cursor
                .pull(64, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES)
                .expect("pull succeeds")
            {
                WorksheetCursorPull::Rows { rows, .. } => {
                    assert!(!rows.is_empty());
                    assert!(rows.len() <= 64);
                    batches.push(rows);
                    inflated_snapshots.push(
                        archive
                            .operation
                            .active()
                            .unwrap()
                            .usage()
                            .unwrap()
                            .operation_inflated_bytes,
                    );
                }
                WorksheetCursorPull::Finished(tail) => {
                    break (batches, tail);
                }
            }
        };

        assert_eq!(batches.iter().map(Vec::len).sum::<usize>(), 600);
        assert_eq!(
            batches
                .iter()
                .flatten()
                .map(|row| row.index)
                .collect::<Vec<_>>(),
            (1..=600).collect::<Vec<_>>()
        );
        assert!(
            inflated_snapshots.windows(2).any(|pair| pair[0] < pair[1]),
            "the same operation continues inflating after earlier row pulls"
        );
        assert!(tail.shell_xml.contains("<sheetData></sheetData>"));
        assert!(archive.operation.active().unwrap().usage().is_some());
        archive.finish_operation().expect("same operation finishes");
    }

    #[test]
    fn late_malformed_tail_never_produces_a_committable_transaction() {
        let xml = worksheet(600, "</sheetData><broken>");
        let mut archive = open_zip(package(&xml)).expect("package opens");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional = Vec::new();
        let mut finished = false;

        loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional.extend(rows),
                Ok(WorksheetCursorPull::Finished(_)) => {
                    finished = true;
                    break;
                }
                Err(error) => {
                    assert!(error.contains("EOF") || error.contains("closed"), "{error}");
                    break;
                }
            }
        }

        assert!(
            !provisional.is_empty(),
            "earlier pulls are intentionally provisional"
        );
        assert!(!finished, "malformed tail must prevent commit");
        archive.cancel_operation();
    }

    #[test]
    fn crc_failure_after_row_pulls_never_produces_finished() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut bytes = package(&xml);
        corrupt_crc_consistently(&mut bytes);
        let mut archive = open_zip(bytes).expect("matching forged metadata passes preflight");
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        let mut provisional_rows = 0;

        let error = loop {
            match cursor.pull(7, WORKSHEET_CURSOR_TARGET_PROJECTED_BYTES) {
                Ok(WorksheetCursorPull::Rows { rows, .. }) => provisional_rows += rows.len(),
                Ok(WorksheetCursorPull::Finished(_)) => panic!("CRC failure must prevent commit"),
                Err(error) => break error,
            }
        };

        assert!(provisional_rows > 0);
        assert!(error.contains("CRC"), "{error}");
        archive.cancel_operation();
    }

    #[test]
    fn close_and_cancel_are_idempotent_and_release_the_entry() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();

        let mut closed = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        closed.close();
        closed.close();
        assert_eq!(closed.pull(1, 1).unwrap_err(), "worksheet cursor is closed");

        let mut canceled = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();
        canceled.cancel();
        canceled.cancel();
        assert_eq!(
            canceled.pull(1, 1).unwrap_err(),
            "worksheet cursor is canceled"
        );
        archive
            .finish_operation()
            .expect("released readers allow finish");
    }

    #[test]
    fn pull_clamps_rows_and_stages_the_first_soft_projection_overrun() {
        let xml = worksheet(600, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        let first = cursor.pull(usize::MAX, usize::MAX).unwrap();
        let WorksheetCursorPull::Rows { rows, .. } = first else {
            panic!("large worksheet must yield rows");
        };
        assert_eq!(rows.len(), WORKSHEET_CURSOR_PULL_ROWS);

        let mut observed = rows.into_iter().map(|row| row.index).collect::<Vec<_>>();
        let mut saw_multi_row_projection_limited_batch = false;
        while let WorksheetCursorPull::Rows {
            rows,
            projected_bytes,
        } = cursor.pull(usize::MAX, 500).unwrap()
        {
            assert!(!rows.is_empty());
            if rows.len() > 1 {
                saw_multi_row_projection_limited_batch = true;
                assert!(projected_bytes <= 500);
            }
            observed.extend(rows.into_iter().map(|row| row.index));
        }
        assert!(saw_multi_row_projection_limited_batch);
        assert_eq!(observed, (1..=600).collect::<Vec<_>>());
        archive.finish_operation().unwrap();
    }

    #[test]
    fn indivisible_row_may_cross_soft_projection_target_but_not_hard_row_cap() {
        let xml = worksheet(2, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        archive.begin_operation("parse-sheet").unwrap();
        let mut cursor = archive
            .open_worksheet_cursor(PART, Rc::from([]), Rc::from([]))
            .unwrap();

        for expected in 1..=2 {
            let WorksheetCursorPull::Rows {
                rows,
                projected_bytes,
            } = cursor.pull(128, 1).unwrap()
            else {
                panic!("row is returned atomically");
            };
            assert_eq!(rows.len(), 1);
            assert_eq!(rows[0].index, expected);
            assert!(projected_bytes > 1);
            assert!(projected_bytes <= crate::worksheet_projector::STREAMED_ROW_PROJECTION_BYTES);
        }
        assert!(matches!(
            cursor.pull(128, 1).unwrap(),
            WorksheetCursorPull::Finished(_)
        ));
        archive.finish_operation().unwrap();
    }

    #[test]
    fn cursor_factory_requires_an_explicit_active_operation() {
        let xml = worksheet(1, "</sheetData></worksheet>");
        let mut archive = open_zip(package(&xml)).unwrap();
        let error = match archive.open_worksheet_cursor(PART, Rc::from([]), Rc::from([])) {
            Ok(_) => panic!("cursor factory must not create a compatibility operation"),
            Err(error) => error,
        };
        assert_eq!(error, "xlsx package operation is not active");
        assert!(!archive.operation.is_active());
    }
}
