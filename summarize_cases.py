from pathlib import Path
from oneill_tracker.pipeline import main

if __name__ == "__main__":
    # Code inside this block only runs when
    # you execute this file directly.
    main(data_dir=Path(__file__).resolve().parent)
