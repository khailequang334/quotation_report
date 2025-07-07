#!/usr/bin/env python3

from quotation_app import QuotationApp


def main():
    try:
        app = QuotationApp()
        app.run()
    except KeyboardInterrupt:
        print("\n\nApplication interrupted by user.")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        print("Please check the log file for more details.")


if __name__ == "__main__":
    main()

