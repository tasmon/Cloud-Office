/* ============================================================
   Cloud Office — generic undo/redo history manager.
   Works on any JSON-serializable state object (deep-cloned snapshots).
   ============================================================ */
class UndoManager {
  constructor(limit = 80) {
    this.limit = limit;
    this.undoStack = [];
    this.redoStack = [];
    this._suspended = false;
  }

  // Call BEFORE mutating `state`, passing the state as it is right now.
  snapshot(state) {
    if (this._suspended) return;
    this.undoStack.push(JSON.stringify(state));
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack.length = 0;
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  // Pass the CURRENT state so it can be pushed onto the redo stack.
  // Returns the parsed previous state, or null if nothing to undo.
  undo(currentState) {
    if (!this.canUndo()) return null;
    this.redoStack.push(JSON.stringify(currentState));
    const prev = this.undoStack.pop();
    return JSON.parse(prev);
  }

  redo(currentState) {
    if (!this.canRedo()) return null;
    this.undoStack.push(JSON.stringify(currentState));
    const next = this.redoStack.pop();
    return JSON.parse(next);
  }

  clear() { this.undoStack.length = 0; this.redoStack.length = 0; }
}
