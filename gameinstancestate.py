class GameInstanceState:
    # Round details
    __round_started_spectation: bool = False

    # Error details
    __error_blank_screen_count: int = 0
    __error_restart_required: bool = False

    def set_round_started_spectation(self, round_started_spectation: bool):
        self.__round_started_spectation = round_started_spectation

    def round_started_spectation(self) -> bool:
        return self.__round_started_spectation

    def increase_error_blank_screen_count(self):
        self.__error_blank_screen_count += 1

    def get_error_blank_screen_count(self) -> int:
        return self.__error_blank_screen_count

    def set_error_restart_required(self, restart_required: bool):
        self.__error_restart_required = restart_required

    def error_restart_required(self) -> bool:
        return self.__error_restart_required

    def reset_error_blank_screen_count(self):
        self.__error_blank_screen_count = 0
