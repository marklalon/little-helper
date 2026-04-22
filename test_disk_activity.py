#!/usr/bin/env python
"""Test disk activity time reading."""

import sys
import os
sys.path.insert(0, 'src')

from system_overlay import init_lhm, get_disk_stats

if __name__ == '__main__':
    if init_lhm():
        stats = get_disk_stats()
        print('Disk stats:')
        print(f'  Temperatures: {stats.get("disk_temps", {})}')
        print(f'  Activity (%): {stats.get("disk_activity", {})}')
    else:
        print('Failed to initialize LHM')
