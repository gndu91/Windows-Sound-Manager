# Code found in the Python Mailing List:
#   https://mail.python.org/pipermail/python-win32/2014-March/013080.html
# Author: Tim Roberts timr at probo.com
# Tested on my computer (Windows 10, Python 3.7, AMD64)
import operator
import random
import unittest
from ctypes.wintypes import BOOL

import comtypes.client
from comtypes import *

try:
    import win32com
except ImportError as e:
    raise ImportError(
        'An error occurred while trying to import some packages.\n'
        'Make sure pypiwin32 is installed (command: `pip install pypiwin32`)\n'
        f'Origin: {e!r}'
    )

MMDeviceApiLib = \
    GUID('{2FDAAFA3-7523-4F66-9957-9D5E7FE698F6}')
IID_IMMDevice = \
    GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
IID_IMMDeviceEnumerator = \
    GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
CLSID_MMDeviceEnumerator = \
    GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')
IID_IMMDeviceCollection = \
    GUID('{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
IID_IAudioEndpointVolume = \
    GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')


class IMMDeviceCollection(IUnknown):
    _iid_ = GUID('{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
    pass


class IAudioEndpointVolume(IUnknown):
    _iid_ = GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')
    _methods_ = [
        STDMETHOD(HRESULT, 'RegisterControlChangeNotify', []),
        STDMETHOD(HRESULT, 'UnregisterControlChangeNotify', []),
        STDMETHOD(HRESULT, 'GetChannelCount', []),
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevel',
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevelScalar',
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevel',
                  (['out', 'retval'], POINTER(c_float), 'pfLevelDB')
                  ),
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevelScalar',
                  (['out', 'retval'], POINTER(c_float), 'pfLevelDB')
                  ),
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevel',
                  (['in'], DWORD, 'nChannel'),
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevelScalar',
                  (['in'], DWORD, 'nChannel'),
                  (['in'], c_float, 'fLevelDB'),
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevel',
                  (['in'], DWORD, 'nChannel'),
                  (['out', 'retval'], POINTER(c_float), 'pfLevelDB')
                  ),
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevelScalar',
                  (['in'], DWORD, 'nChannel'),
                  (['out', 'retval'], POINTER(c_float), 'pfLevelDB')
                  ),
        COMMETHOD([], HRESULT, 'SetMute',
                  (['in'], BOOL, 'bMute'),
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'GetMute',
                  (['out', 'retval'], POINTER(BOOL), 'pbMute')
                  ),
        COMMETHOD([], HRESULT, 'GetVolumeStepInfo',
                  (['out', 'retval'], POINTER(c_float), 'pnStep'),
                  (['out', 'retval'], POINTER(c_float), 'pnStepCount'),
                  ),
        COMMETHOD([], HRESULT, 'VolumeStepUp',
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'VolumeStepDown',
                  (['in'], POINTER(GUID), 'pguidEventContext')
                  ),
        COMMETHOD([], HRESULT, 'QueryHardwareSupport',
                  (['out', 'retval'], POINTER(DWORD), 'pdwHardwareSupportMask')
                  ),
        COMMETHOD([], HRESULT, 'GetVolumeRange',
                  (['out', 'retval'], POINTER(c_float), 'pfMin'),
                  (['out', 'retval'], POINTER(c_float), 'pfMax'),
                  (['out', 'retval'], POINTER(c_float), 'pfIncr')
                  ),

    ]


class IMMDevice(IUnknown):
    _iid_ = GUID('{D666063F-1587-4E43-81F1-B948E807363F}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'Activate',
                  (['in'], POINTER(GUID), 'iid'),
                  (['in'], DWORD, 'dwClsCtx'),
                  (['in'], POINTER(DWORD), 'pActivationParans'),
                  (['out', 'retval'], POINTER(POINTER(IAudioEndpointVolume)), 'ppInterface')
                  ),
        STDMETHOD(HRESULT, 'OpenPropertyStore', []),
        STDMETHOD(HRESULT, 'GetId', []),
        STDMETHOD(HRESULT, 'GetState', [])
    ]
    pass


class IMMDeviceEnumerator(comtypes.IUnknown):
    _iid_ = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')

    _methods_ = [
        COMMETHOD([], HRESULT, 'EnumAudioEndpoints',
                  (['in'], DWORD, 'dataFlow'),
                  (['in'], DWORD, 'dwStateMask'),
                  (['out', 'retval'], POINTER(POINTER(IMMDeviceCollection)), 'ppDevices')
                  ),
        COMMETHOD([], HRESULT, 'GetDefaultAudioEndpoint',
                  (['in'], DWORD, 'dataFlow'),
                  (['in'], DWORD, 'role'),
                  (['out', 'retval'], POINTER(POINTER(IMMDevice)), 'ppDevices')
                  )
    ]


class AudioDevice:
    def __init__(self, dev):
        self.__device = dev

    @staticmethod
    def get_default_output_device():
        """Return the default device.
        Warning: the device may not always be the default one, for
            instance, if the user connects/disconnects earphones."""
        enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER
        )
        endpoint = enumerator.GetDefaultAudioEndpoint(0, 1)
        return AudioDevice(endpoint.Activate(IID_IAudioEndpointVolume, comtypes.CLSCTX_INPROC_SERVER, None))

    @classmethod
    def sanitize_volume(cls, volume):
        volume = float(volume)
        if volume < 0:
            return 0
        if volume > 100:
            return 100
        return volume

    @property
    def current(self) -> float:
        """Return the current in percent."""
        return round(self.__device.GetMasterVolumeLevelScalar() * 100)

    @current.setter
    def current(self, value: float):
        # TODO: Learn more about the `pguidEventContext` argument
        print('Set to', self.sanitize_volume(value))
        self.__device.SetMasterVolumeLevelScalar(self.sanitize_volume(value) / 100, None)

    @property
    def mute(self):
        return bool(self.__device.GetMute())

    @mute.setter
    def mute(self, value):
        self.__device.SetMute(bool(value), None)

    def increase(self):
        self.__device.VolumeStepUp(None)

    def decrease(self):
        self.__device.VolumeStepDown(None)

    def __iadd__(self, other):
        self.current = self.current + other
        return self

    def __isub__(self, other):
        self.current = self.current - other
        return self

    def __imul__(self, other):
        self.current = self.current * other
        return self

    def __idiv__(self, other):
        self.current = self.current / other
        return self

    def __imod__(self, other):
        self.current = self.current % other
        return self

    # TODO: Add rounding before we compare
    def __cmp__(self, other):
        return (self.current > other) - (other > self.current)

    def __gt__(self, other):
        return self.__cmp__(other) > 0

    def __ge__(self, other):
        return self.__cmp__(other) >= 0

    def __eq__(self, other):
        return self.__cmp__(other) == 0

    def __ne__(self, other):
        return self.__cmp__(other) != 0

    def __le__(self, other):
        return self.__cmp__(other) <= 0

    def __lt__(self, other):
        return self.__cmp__(other) < 0

    def __int__(self):
        return int(self.current)

    def __float__(self):
        return self.current

    def __bool__(self):
        return bool(int(self))

    def __add__(self, other):
        return self.current + other

    def __sub__(self, other):
        return self.current - other

    def __mul__(self, other):
        return self.current * other

    def __truediv__(self, other):
        return operator.truediv(self.current, other)

    def __divmod__(self, other):
        return divmod(self.current, other)

    def __mod__(self, other):
        return self.current % other

    def __neg__(self):
        return -self.current

    def __abs__(self):
        """The volume is always positive"""
        return self.current

    def __radd__(self, other):
        return self.current + other

    def __rdiv__(self, other):
        return other / self.current

    def __rdivmod__(self, other):
        return divmod(other, self.current)

    def __rfloordiv__(self, other):
        return other // self.current

    def __rmod__(self, other):
        return operator.mod(other, self.current)

    def __rmul__(self, other):
        return self.current * other

    def __rsub__(self, other):
        return other - self.current


__device = None


def refresh_device():
    """Refresh the current audio device."""
    global __device
    __device = AudioDevice.get_default_output_device()


def get_volume_device() -> AudioDevice:
    """
    TODO: Force the device to be refreshed in a regular basis, for instance
        in the case a headphone is connected/disconnected
    TODO: Decide whether or not race conditions can be a problem
    :return:
    """
    if __device is None:
        refresh_device()
    assert isinstance(__device, AudioDevice)
    return __device


class Tests(unittest.TestCase):
    # TODO: Add more tests
    def test_operations(self):
        device = get_volume_device()
        assert 1 + device - 1 == device
        old = device.current
        device *= 1
        self.assertEqual(device, old)
        device /= 1
        self.assertEqual(device, old)
        device += 0
        self.assertEqual(device, old)
        device -= 0
        self.assertEqual(device, old)
        self.assertEqual(device - device, 0)
        self.assertEqual(device + device, device * 2)
        self.assertEqual(device + device, 2 * device)
        self.assertEqual(device % 100, old)
        for i in range(256):
            string = '\t'.join((
                random.choice(('device', f'{random.random() * 1024}')),
                random.choice('+*/-%'),
                random.choice(('device', f'{(random.random()) * 1024}'))))
            print(string + ' = ' + str(eval(string)))
            self.assertEqual(eval(string), eval(string.replace('device', str(float(device)))))

    def test_setter(self):
        device = get_volume_device()
        old = device.current
        assert old == device.current
        device.current = old
        assert old == device.current
        assert old == device
        assert old + 1 == device + 1
        assert old + 1 != device - 1

        print(('%r' % device.current) + '%')
        device.current = 64
        device.mute = 0
        device += 1
        print(('%r' % device.current) + '%')
        device += 1
        print(repr(device.mute))
