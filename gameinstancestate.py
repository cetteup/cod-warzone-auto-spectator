class GameInstanceState:
    __round_rotation_started: bool = False
    __searching_for_game: bool = False
    __in_pre_game: bool = False

    # Round details
    __round_started_spectation: bool = False

    # Error details
    __error_blank_screen_count: int = 0
    __error_restart_required: bool = False

    def set_round_rotation_started(self, round_rotation_started: bool):
        self.__round_rotation_started = round_rotation_started

    def round_rotation_started(self) -> bool:
        return self.__round_rotation_started

    def set_searching_for_game(self, searching_for_game: bool):
        self.__searching_for_game = searching_for_game

    def searching_for_game(self) -> bool:
        return self.__searching_for_game

    def set_in_pre_game(self, in_pre_game: bool):
        self.__in_pre_game = in_pre_game

    def in_pre_game(self) -> bool:
        return self.__in_pre_game

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

    def round_end_reset(self):
        self.__round_rotation_started = False
        self.__searching_for_game = False
        self.__in_pre_game = False
        self.__round_started_spectation = False
