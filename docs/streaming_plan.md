# Streaming/Incremental Output Implementation Plan

## Problem Statement
The current Excel processor loads entire sheets into memory before processing and only writes output after complete processing. With large Excel files, this can lead to:
- Memory issues or crashes
- Complete data loss on errors
- Long waiting times without partial results

## Solution Overview
Implement a streaming approach that processes and saves data incrementally, allowing for:
- Reduced memory footprint
- Partial results preservation on errors
- Resumable processing
- Progress visibility during long operations

## Implementation Roadmap

### Phase 1: Chunked Data Extraction
**Files to modify:**
- `core/extractor.py`

**Changes required:**
1. Refactor `DataExtractor.extract_data()` to support yielding chunks of processed rows
2. Add memory monitoring to dynamically adjust chunk sizes
3. Preserve extraction state between chunks to ensure consistency

### Phase 2: Progressive Output Writing
**Files to modify:**
- `output/formatter.py`
- `output/writer.py`

**Changes required:**
1. Add `format_chunk()` method to `OutputFormatter` for partial results
2. Create `StreamingOutputWriter` with JSON Lines support
3. Implement append operations for existing output files
4. Add temporary file management for partial results

### Phase 3: Checkpoint System
**Files to modify:**
- `utils/caching.py`
- New file: `utils/checkpointing.py`

**Changes required:**
1. Create checkpoint data structure to track processing state
2. Implement checkpoint writing/reading functionality
3. Add integrity verification for checkpoints
4. Implement cleanup for completed checkpoints

### Phase 4: Workflow Integration
**Files to modify:**
- `workflows/base_workflow.py`
- `workflows/multi_sheet.py`
- `workflows/batch.py`
- `workflows/single_file.py`

**Changes required:**
1. Add streaming mode to all workflow classes
2. Implement sheet-level checkpoint management
3. Add error recovery with partial result preservation
4. Modify progress reporting for chunk-based updates

### Phase 5: Configuration and CLI
**Files to modify:**
- `config.py`
- `cli.py`

**Changes required:**
1. Add streaming mode configuration options:
   - `--streaming` flag
   - `--chunk-size` parameter
   - `--checkpoint-dir` parameter
2. Add resume functionality:
   - `--resume` flag
   - `--checkpoint-id` parameter
3. Update help documentation

## Implementation Details

### Checkpoint Data Structure
```json
{
  "checkpoint_id": "unique-id",
  "file": "path/to/excel.xlsx",
  "timestamp": "2023-06-15T14:30:00",
  "state": {
    "current_sheet": "Sheet1",
    "sheets_completed": ["Sheet2", "Sheet3"],
    "current_chunk": 5,
    "total_chunks_estimated": 20,
    "rows_processed": 500,
    "output_files": {
      "Sheet2": "temp/Sheet2_complete.jsonl",
      "Sheet3": "temp/Sheet3_complete.jsonl",
      "Sheet1": "temp/Sheet1_partial.jsonl"
    }
  }
}
```

### Streaming Workflow Logic
1. Check for existing checkpoint
2. If resuming, restore state from checkpoint
3. Process each sheet:
   - For each chunk of rows:
     - Extract and process chunk
     - Write chunk to temporary file
     - Update checkpoint
   - After sheet completion, finalize sheet output
4. After all sheets, combine outputs
5. Clean up temporary files and checkpoints

## Testing Plan
1. Unit tests for chunked extraction
2. Integration tests for checkpoint recovery
3. Performance tests with large Excel files
4. Error injection tests to verify recovery

## Success Metrics
1. Successfully process Excel files exceeding available RAM
2. 100% data preservation on interruption/error
3. Resume functionality works across process restarts
4. Memory usage remains stable regardless of input size

## Timeline
- Phase 1: 3 days
- Phase 2: 2 days
- Phase 3: 2 days
- Phase 4: 3 days
- Phase 5: 1 day
- Testing & Refinement: 3 days

Total: ~2 weeks
