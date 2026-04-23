import { Workbook } from "../../models/Workbook";

export interface WorkbookSnapshot {
  entries: Map<string, Uint8Array>;
  sheetPathByName: Map<string, string>;
  sheetXmlByName: Map<string, string>;
}

const _snapshotByWorkbook = new WeakMap<Workbook, WorkbookSnapshot>();

export function setWorkbookSnapshot(workbook: Workbook, snapshot: WorkbookSnapshot): void {
  _snapshotByWorkbook.set(workbook, snapshot);
}

export function getWorkbookSnapshot(workbook: Workbook): WorkbookSnapshot | undefined {
  return _snapshotByWorkbook.get(workbook);
}
