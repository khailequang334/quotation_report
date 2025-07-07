import sys


class ProgressTracker:
    
    def __init__(self):
        self.bar_length = 100
    
    def print_progress_bar(self, progress: float) -> None:
        filled_length = int(self.bar_length * progress)
        bar = '[' + '#' * filled_length + ' ' * (self.bar_length - filled_length) + ']'
        sys.stdout.write('\r' + bar + ' %d%%' % (progress * 100))
        sys.stdout.flush()
    
    def update_progress(self, current: int, total: int) -> None:
        if total > 0:
            self.print_progress_bar(current / total) 