import functools
from datetime import date
from sound import Sound

if __name__ == '__main__':

    def switch(value, cases, default=None):
        for i, j in cases:
            if i == value:
                return j()
        if default is not None:
            return default(value)


    print(
        ""
        "Windows Sound Manager (Python 3)\n"
        f"Copyright (c) 2014 - {date.today().year} \n"
        "Software provided as it, without any guarantee"
    )

    while True:
        print(
            "Choose an option:\n"
            "1] Mute / Unmute volume\n"
            "2] Increase volume\n"
            "3] Decrease volume\n"
            "4] Set volume to 0\n"
            "5] Set volume to 100\n"
            "6] Set volume to ...\n"
            "7] Print sound settings\n"
            "8] Quit"
        )
        option = input("> ")
        print("")

        switch(option, (
            ('1', Sound.mute),
            ('2', Sound.volume_up),
            ('3', Sound.volume_down),
            ('4', Sound.volume_min),
            ('5', Sound.volume_max),
            ('6', (lambda: Sound.volume_set(int(input("Volume (0 - 100): "))))),
            ('7', (lambda: print(
                f"Current volume | {str(Sound.current_volume())}\n"
                f"Sound muted    | {str(Sound.is_muted())}\n"
                "----------------------\n"
            ))),
            ('8', functools.partial(exit, 0))
        ), (lambda v: print(f'Unknown command ({v!r}), please retry')))
