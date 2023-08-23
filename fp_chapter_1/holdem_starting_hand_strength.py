import collections

holdem_player_hand = collections.namedtuple('HPCards', ['rank', 'paired', 'suited', 'combos', 'combo_percent', 'win_percent'])

def suitedvalues(stringrank):
    hmh = holdem_player_hand(
            rank=0, paired=False, suited=True, 
            combos=4, combo_percent=.30, win_percent=0)
    if stringrank == 'AK':
        hmh = hmh._replace(rank=1, win_percent=22)
    elif stringrank == 'AQ':
        return 2
    elif stringrank == 'AJ':
        return 3
    elif stringrank == 'AT':
        return 4
    elif stringrank == 'A9':
        return 5
    elif stringrank == 'A8':
        return 6
    elif stringrank == 'A7':
        return 7
    elif stringrank == 'A6':
        return 8
    elif stringrank == 'A5':
        return 9
    elif stringrank == 'A4':
        return 10
    elif stringrank == 'A3':
        return 11
    elif stringrank == 'A2':
        return 12
    elif stringrank == 'KQ':
        return 13
    elif stringrank == 'KJ':
        return 14
    elif stringrank == 'KT':
        return 15
    elif stringrank == 'K9':
        return 16
    elif stringrank == 'K8':
        return 17
    elif stringrank == 'K7':
        return 18
    elif stringrank == 'K6':
        return 19
    elif stringrank == 'K5':
        return 20
    elif stringrank == 'K4':
        return 21
    elif stringrank == 'K3':
        return 22
    elif stringrank == 'K2':
        return 23
    elif stringrank == 'QJ':
        return 24
    elif stringrank == 'QT':
        return 25
    elif stringrank == 'Q9':
        return 26
    elif stringrank == 'Q8':
        return 27
    elif stringrank == 'Q7':
        return 28
    elif stringrank == 'Q6':
        return 29
    elif stringrank == 'Q5':
        return 30
    elif stringrank == 'Q4':
        return 31
    elif stringrank == 'Q3':
        return 32
    elif stringrank == 'Q2':
        return 33
    elif stringrank == 'JT':
        return 34
    elif stringrank == 'J9':
        return 35
    elif stringrank == 'J8':
        return 36
    elif stringrank == 'J7':
        return 37
    elif stringrank == 'J6':
        return 38
    elif stringrank == 'J5':
        return 39
    elif stringrank == 'J4':
        return 40
    elif stringrank == 'J3':
        return 41
    elif stringrank == 'J2':
        return 42
    elif stringrank == 'T9':
        return 43
    elif stringrank == 'T8':
        return 44
    elif stringrank == 'T7':
        return 45
    elif stringrank == 'T6':
        return 46
    elif stringrank == 'T5':
        return 47
    elif stringrank == 'T4':
        return 48
    elif stringrank == 'T3':
        return 49
    elif stringrank == 'T2':
        return 50
    elif stringrank == '98':
        return 51
    elif stringrank == '97':
        return 52
    elif stringrank == '96':
        return 53
    elif stringrank == '95':
        return 54
    elif stringrank == '94':
        return 55
    elif stringrank == '93':
        return 56
    elif stringrank == '92':
        return 57
    elif stringrank == '87':
        return 58
    elif stringrank == '86':
        return 59
    elif stringrank == '85':
        return 60
    elif stringrank == '84':
        return 61
    elif stringrank == '83':
        return 62
    elif stringrank == '82':
        return 63
    elif stringrank == '76':
        return 64
    elif stringrank == '75':
        return 65
    elif stringrank == '74':
        return 66
    elif stringrank == '73':
        return 67
    elif stringrank == '72':
        return 68
    elif stringrank == '65':
        return 69
    elif stringrank == '64':
        return 70
    elif stringrank == '63':
        return 71
    elif stringrank == '62':
        return 72
    elif stringrank == '54':
        return 73
    elif stringrank == '53':
        return 74
    elif stringrank == '52':
        return 75
    elif stringrank == '43':
        return 76
    elif stringrank == '42':
        return 77
    elif stringrank == '32':
        return 78
    else:
        return 0
    return hmh

def unsuitedvalues(stringrank):
    if stringrank == 'AK':
        return 101
    elif stringrank == 'AQ':
        return 102
    elif stringrank == 'AJ':
        return 103
    elif stringrank == 'AT':
        return 104
    elif stringrank == 'A9':
        return 105
    elif stringrank == 'A8':
        return 106
    elif stringrank == 'A7':
        return 107
    elif stringrank == 'A6':
        return 108
    elif stringrank == 'A5':
        return 109
    elif stringrank == 'A4':
        return 110
    elif stringrank == 'A3':
        return 111
    elif stringrank == 'A2':
        return 112
    elif stringrank == 'KQ':
        return 113
    elif stringrank == 'KJ':
        return 114
    elif stringrank == 'KT':
        return 115
    elif stringrank == 'K9':
        return 116
    elif stringrank == 'K8':
        return 117
    elif stringrank == 'K7':
        return 118
    elif stringrank == 'K6':
        return 119
    elif stringrank == 'K5':
        return 120
    elif stringrank == 'K4':
        return 121
    elif stringrank == 'K3':
        return 122
    elif stringrank == 'K2':
        return 123
    elif stringrank == 'QJ':
        return 124
    elif stringrank == 'QT':
        return 125
    elif stringrank == 'Q9':
        return 126
    elif stringrank == 'Q8':
        return 127
    elif stringrank == 'Q7':
        return 128
    elif stringrank == 'Q6':
        return 129
    elif stringrank == 'Q5':
        return 130
    elif stringrank == 'Q4':
        return 131
    elif stringrank == 'Q3':
        return 132
    elif stringrank == 'Q2':
        return 133
    elif stringrank == 'JT':
        return 134
    elif stringrank == 'J9':
        return 135
    elif stringrank == 'J8':
        return 136
    elif stringrank == 'J7':
        return 137
    elif stringrank == 'J6':
        return 138
    elif stringrank == 'J5':
        return 139
    elif stringrank == 'J4':
        return 140
    elif stringrank == 'J3':
        return 141
    elif stringrank == 'J2':
        return 142
    elif stringrank == 'T9':
        return 143
    elif stringrank == 'T8':
        return 144
    elif stringrank == 'T7':
        return 145
    elif stringrank == 'T6':
        return 146
    elif stringrank == 'T5':
        return 147
    elif stringrank == 'T4':
        return 148
    elif stringrank == 'T3':
        return 149
    elif stringrank == 'T2':
        return 150
    elif stringrank == '98':
        return 151
    elif stringrank == '97':
        return 152
    elif stringrank == '96':
        return 153
    elif stringrank == '95':
        return 154
    elif stringrank == '94':
        return 155
    elif stringrank == '93':
        return 156
    elif stringrank == '92':
        return 157
    elif stringrank == '87':
        return 158
    elif stringrank == '86':
        return 159
    elif stringrank == '85':
        return 160
    elif stringrank == '84':
        return 161
    elif stringrank == '83':
        return 162
    elif stringrank == '82':
        return 163
    elif stringrank == '76':
        return 164
    elif stringrank == '75':
        return 165
    elif stringrank == '74':
        return 166
    elif stringrank == '73':
        return 167
    elif stringrank == '72':
        return 168
    elif stringrank == '65':
        return 169
    elif stringrank == '64':
        return 170
    elif stringrank == '63':
        return 171
    elif stringrank == '62':
        return 172
    elif stringrank == '54':
        return 173
    elif stringrank == '53':
        return 174
    elif stringrank == '52':
        return 175
    elif stringrank == '43':
        return 176
    elif stringrank == '42':
        return 177
    elif stringrank == '32':
        return 178
    else:
        return 0
    
def pairedvalues(stringrank):
    hmh = holdem_player_hand(
            paired=True, suited=False, 
            combos=6, combo_percent=.45)
    if stringrank == 'AA':
        return hmh(rank=1, win_percent=22)
    elif stringrank == 'KK':
        return hmh(rank=2, win_percent=22)
    else:
        return 0