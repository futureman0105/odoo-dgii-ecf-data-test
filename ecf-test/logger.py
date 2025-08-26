import logging
import sys

class Logger:
    def __init__(self, name="MyLogger", log_to_console=True, log_to_file=None, level=logging.DEBUG):
        """
        Custom logger class.

        :param name: Logger name
        :param log_to_console: Print logs to console
        :param log_to_file: If given, path to log file
        :param level: Logging level
        """
        self.logger = logging.getLogger(name)
        self.logger.setLevel(level)
        self.logger.propagate = False  # Avoid duplicate logs if root logger exists

        # Remove existing handlers
        if self.logger.hasHandlers():
            self.logger.handlers.clear()

        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', '%Y-%m-%d %H:%M:%S')

        # Console handler
        if log_to_console:
            ch = logging.StreamHandler(sys.stdout)
            ch.setFormatter(formatter)
            self.logger.addHandler(ch)

        # File handler
        if log_to_file:
            fh = logging.FileHandler(log_to_file)
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)

    # Logging methods
    def debug(self, msg):
        self.logger.debug(msg)

    def info(self, msg):
        self.logger.info(msg)

    def warning(self, msg):
        self.logger.warning(msg)

    def error(self, msg):
        self.logger.error(msg)

    def critical(self, msg):
        self.logger.critical(msg)


__logger = Logger("app.log")
