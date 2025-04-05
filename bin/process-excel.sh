#!/bin/bash

# Default configuration values
INPUT_DIR="data/input"
OUTPUT_DIR="data/output/batch_$(date +%Y%m%d_%H%M%S)"
CONFIG_FILE="config/streaming-defaults.json"
SCRIPT_DIR="$(dirname "$(readlink -f "$0")")"
PROJECT_ROOT="$(dirname "$SCRIPT_DIR")"

# Function to display help
show_help() {
    echo "Excel Processor - Batch Processing Script"
    echo ""
    echo "Usage: $0 [options]"
    echo ""
    echo "Options:"
    echo "  -i, --input DIR       Input directory (default: $INPUT_DIR)"
    echo "  -o, --output DIR      Output directory (default: includes timestamp)"
    echo "  -c, --config FILE     Configuration file (default: $CONFIG_FILE)"
    echo "  -h, --help            Show this help message"
    echo ""
    echo "Any additional options will be passed directly to the Excel processor."
    echo "Example: $0 --log-level debug"
    exit 0
}

# Parse command-line arguments
ADDITIONAL_ARGS=""
while [[ $# -gt 0 ]]; do
    case $1 in
        -i|--input)
            INPUT_DIR="$2"
            shift 2
            ;;
        -o|--output)
            OUTPUT_DIR="$2"
            shift 2
            ;;
        -c|--config)
            CONFIG_FILE="$2"
            shift 2
            ;;
        -h|--help)
            show_help
            ;;
        *)
            ADDITIONAL_ARGS="$ADDITIONAL_ARGS $1"
            shift
            ;;
    esac
done

# Activate virtual environment if it exists
VENV_PATH="$PROJECT_ROOT/.venv"
if [ -f "$VENV_PATH/bin/activate" ]; then
    source "$VENV_PATH/bin/activate"
fi

# Change to project root
cd "$PROJECT_ROOT"

# Run the processor with the specified options
echo "Running batch processing with the following settings:"
echo "Input directory: $INPUT_DIR"
echo "Output directory: $OUTPUT_DIR"
echo "Configuration file: $CONFIG_FILE"
echo "Additional arguments: $ADDITIONAL_ARGS"
echo ""

python cli.py batch -i "$INPUT_DIR" -o "$OUTPUT_DIR" --config "$CONFIG_FILE" $ADDITIONAL_ARGS

# Check the exit status
EXIT_STATUS=$?
if [ $EXIT_STATUS -eq 0 ]; then
    echo ""
    echo "Batch processing completed successfully."
    echo "Output saved to: $OUTPUT_DIR"
else
    echo ""
    echo "Batch processing failed with exit code $EXIT_STATUS."
fi

exit $EXIT_STATUS
