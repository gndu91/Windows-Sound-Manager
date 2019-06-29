# import keyboard as Keyboard
# Keyboard.key = Keyboard.press_and_release
import os

if os.name == 'nt':
    from _nt import get_volume_device
else:
    raise NotImplementedError('Only available in Windows for the moment')


class Sound:
    """
    Class Sound
    :description: Based on the class written by Paradoxis <luke@paradoxis.nl>, but using the Windows API
    """

    @staticmethod
    def current_volume():
        """
        Current volume getter
        :return: int
        """
        return get_volume_device().current

    @staticmethod
    def __set_current_volume(volume):
        """
        Current volume setter
        Volume is sanitized in the device setter
        :return: void
        """
        get_volume_device().current = volume

    # The sound is not muted by default, better tracking should be made
    __is_muted = False

    @staticmethod
    def is_muted():
        """
        Is muted getter
        :return: boolean
        """
        return get_volume_device().mute

    @staticmethod
    def __track():
        """
        Start tracking the sound and mute settings
        :return: void
        """
        pass

    @staticmethod
    def mute():
        """
        Mute or un-mute the system sounds
        Done by triggering a fake VK_VOLUME_MUTE key event
        :return: void
        """
        get_volume_device().mute = not get_volume_device().mute

    @staticmethod
    def volume_up():
        """
        Increase system volume
        Done by triggering a fake VK_VOLUME_UP key event
        :return: void
        """
        get_volume_device().increase()

    @staticmethod
    def volume_down():
        """
        Decrease system volume
        Done by triggering a fake VK_VOLUME_DOWN key event
        :return: void
        """
        get_volume_device().decrease()

    @staticmethod
    def volume_set(amount):
        """
        Set the volume to a specific volume, limited to even numbers.
        This is due to the fact that a VK_VOLUME_UP/VK_VOLUME_DOWN event increases
        or decreases the volume by two every single time.
        :return: void
        """
        get_volume_device().current = amount

    @staticmethod
    def volume_min():
        """
        Set the volume to min (0)
        :return: void
        """
        get_volume_device().current = 0

    @staticmethod
    def volume_max():
        """
        Set the volume to max (100)
        :return: void
        """
        get_volume_device().current = 100
