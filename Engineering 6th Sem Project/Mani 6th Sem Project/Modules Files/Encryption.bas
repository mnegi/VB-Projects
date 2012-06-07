Attribute VB_Name = "Encryption"
'****************************************************
'   CODED BY: MANOHAR SINGH NEGI                    *
'             6th Semester , I.S.E.                 *
'             R.V. College Of Engineering           *
'             Bangalore - 560059                    *
'             manohar.negi@gmail.com                *
'                                                   *
'****************************************************

Public Function Encrypt(Key As Integer) As Integer
Select Case Key
Case 33
Encrypt = 90
Case 34
Encrypt = 66
Case 35
Encrypt = 87
Case 36
Encrypt = 78
Case 37
Encrypt = 82
Case 38
Encrypt = 76
Case 39
Encrypt = 80
Case 40
Encrypt = 85
Case 41
Encrypt = 89
Case 42
Encrypt = 67
Case 43
Encrypt = 72
Case 44
Encrypt = 74
Case 45
Encrypt = 77
Case 46
Encrypt = 79
Case 47
Encrypt = 83
Case 48
Encrypt = 84
Case 49
Encrypt = 86
Case 50
Encrypt = 65
Case 51
Encrypt = 68
Case 52
Encrypt = 71
Case 53
Encrypt = 73
Case 54
Encrypt = 75
Case 55
Encrypt = 81
Case 56
Encrypt = 88
Case 57
Encrypt = 70
Case 58
Encrypt = 69
Case 59
Encrypt = 122
Case 60
Encrypt = 98
Case 61
Encrypt = 119
Case 62
Encrypt = 110
Case 63
Encrypt = 114
Case 64
Encrypt = 108
Case 65
Encrypt = 112
Case 66
Encrypt = 117
Case 67
Encrypt = 121
Case 68
Encrypt = 99
Case 69
Encrypt = 104
Case 70
Encrypt = 106
Case 71
Encrypt = 109
Case 72
Encrypt = 111
Case 73
Encrypt = 115
Case 74
Encrypt = 116
Case 75
Encrypt = 118
Case 76
Encrypt = 97
Case 77
Encrypt = 100
Case 78
Encrypt = 103
Case 79
Encrypt = 113
Case 80
Encrypt = 107
Case 81
Encrypt = 105
Case 82
Encrypt = 120
Case 83
Encrypt = 102
Case 84
Encrypt = 101
Case 85
Encrypt = 126
Case 86
Encrypt = 124
Case 87
Encrypt = 46
Case 88
Encrypt = 48
Case 89
Encrypt = 51
Case 90
Encrypt = 57
Case 91
Encrypt = 49
Case 92
Encrypt = 52
Case 93
Encrypt = 50
Case 94
Encrypt = 53
Case 95
Encrypt = 55
Case 96
Encrypt = 54
Case 97
Encrypt = 56
Case 98
Encrypt = 33
Case 99
Encrypt = 92
Case 100
Encrypt = 58
Case 101
Encrypt = 123
Case 102
Encrypt = 93
Case 103
Encrypt = 125
Case 104
Encrypt = 59
Case 105
Encrypt = 61
Case 106
Encrypt = 91
Case 107
Encrypt = 47
Case 108
Encrypt = 96
Case 109
Encrypt = 60
Case 110
Encrypt = 64
Case 111
Encrypt = 42
Case 112
Encrypt = 62
Case 113
Encrypt = 36
Case 114
Encrypt = 43
Case 115
Encrypt = 35
Case 116
Encrypt = 45
Case 117
Encrypt = 39
Case 118
Encrypt = 34
Case 119
Encrypt = 95
Case 120
Encrypt = 40
Case 121
Encrypt = 63
Case 122
Encrypt = 37
Case 123
Encrypt = 44
Case 124
Encrypt = 38
Case 125
Encrypt = 41
Case 126
Encrypt = 94
End Select
End Function

Public Function Decrypt(Key As Integer) As Integer
Select Case Key
Case 90
Decrypt = 33
Case 66
Decrypt = 34
Case 87
Decrypt = 35
Case 78
Decrypt = 36
Case 82
Decrypt = 37
Case 76
Decrypt = 38
Case 80
Decrypt = 39
Case 85
Decrypt = 40
Case 89
Decrypt = 41
Case 67
Decrypt = 42
Case 72
Decrypt = 43
Case 74
Decrypt = 44
Case 77
Decrypt = 45
Case 79
Decrypt = 46
Case 83
Decrypt = 47
Case 84
Decrypt = 48
Case 86
Decrypt = 49
Case 65
Decrypt = 50
Case 68
Decrypt = 51
Case 71
Decrypt = 52
Case 73
Decrypt = 53
Case 75
Decrypt = 54
Case 81
Decrypt = 55
Case 88
Decrypt = 56
Case 70
Decrypt = 57
Case 69
Decrypt = 58
Case 122
Decrypt = 59
Case 98
Decrypt = 60
Case 119
Decrypt = 61
Case 110
Decrypt = 62
Case 114
Decrypt = 63
Case 108
Decrypt = 64
Case 112
Decrypt = 65
Case 117
Decrypt = 66
Case 121
Decrypt = 67
Case 99
Decrypt = 68
Case 104
Decrypt = 69
Case 106
Decrypt = 70
Case 109
Decrypt = 71
Case 111
Decrypt = 72
Case 115
Decrypt = 73
Case 116
Decrypt = 74
Case 118
Decrypt = 75
Case 97
Decrypt = 76
Case 100
Decrypt = 77
Case 103
Decrypt = 78
Case 113
Decrypt = 79
Case 107
Decrypt = 80
Case 105
Decrypt = 81
Case 120
Decrypt = 82
Case 102
Decrypt = 83
Case 101
Decrypt = 84
Case 126
Decrypt = 85
Case 124
Decrypt = 86
Case 46
Decrypt = 87
Case 48
Decrypt = 88
Case 51
Decrypt = 89
Case 57
Decrypt = 90
Case 49
Decrypt = 91
Case 52
Decrypt = 92
Case 50
Decrypt = 93
Case 53
Decrypt = 94
Case 55
Decrypt = 95
Case 54
Decrypt = 96
Case 56
Decrypt = 97
Case 33
Decrypt = 98
Case 92
Decrypt = 99
Case 58
Decrypt = 100
Case 123
Decrypt = 101
Case 93
Decrypt = 102
Case 125
Decrypt = 103
Case 59
Decrypt = 104
Case 61
Decrypt = 105
Case 91
Decrypt = 106
Case 47
Decrypt = 107
Case 96
Decrypt = 108
Case 60
Decrypt = 109
Case 64
Decrypt = 110
Case 42
Decrypt = 111
Case 62
Decrypt = 112
Case 36
Decrypt = 113
Case 43
Decrypt = 114
Case 35
Decrypt = 115
Case 45
Decrypt = 116
Case 39
Decrypt = 117
Case 34
Decrypt = 118
Case 95
Decrypt = 119
Case 40
Decrypt = 120
Case 63
Decrypt = 121
Case 37
Decrypt = 122
Case 44
Decrypt = 123
Case 38
Decrypt = 124
Case 41
Decrypt = 125
Case 94
Decrypt = 126
End Select
End Function

