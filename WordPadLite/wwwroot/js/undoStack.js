// undoStack.js
// WordPad Lite — 1000 ステップ Undo/Redo スナップショット スタック

export class UndoStack {
  constructor(limit = 1000) {
    this.limit = limit;
    this.undoStack = [];
    this.redoStack = [];
  }

  push(snapshot) {
    if (this.undoStack.length && this.undoStack[this.undoStack.length - 1].html === snapshot.html) return;

    this.undoStack.push(snapshot);
    if (this.undoStack.length > this.limit) {
      this.undoStack.shift();
    }
    this.redoStack.length = 0;
  }

  replaceTop(snapshot) {
    if (!this.undoStack.length) this.undoStack.push(snapshot);
    else this.undoStack[this.undoStack.length - 1] = snapshot;
    this.redoStack.length = 0;
  }

  reset(snapshot) {
    this.undoStack = [snapshot];
    this.redoStack.length = 0;
  }

  get canUndo() {
    return this.undoStack.length > 1;
  }

  get canRedo() {
    return this.redoStack.length > 0;
  }

  undo(apply) {
    if (!this.canUndo) return false;
    const current = this.undoStack.pop();
    this.redoStack.push(current);
    apply(this.undoStack[this.undoStack.length - 1]);
    return true;
  }

  redo(apply) {
    if (!this.canRedo) return false;
    const next = this.redoStack.pop();
    this.undoStack.push(next);
    apply(next);
    return true;
  }
}
