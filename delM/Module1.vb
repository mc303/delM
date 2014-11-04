Imports System.IO
Imports System.Runtime.InteropServices
Imports System
Imports System.Collections.Generic
Imports Microsoft.Win32.SafeHandles


Module Module1
    Public Const MAX_PATH As Integer = 260
    Public Const MAX_ALTERNATE As Integer = 14

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure FILETIME
        Public dwLowDateTime As UInteger
        Public dwHighDateTime As UInteger
    End Structure

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)> _
    Public Structure WIN32_FIND_DATA
        Public dwFileAttributes As FileAttributes
        Public ftCreationTime As FILETIME
        Public ftLastAccessTime As FILETIME
        Public ftLastWriteTime As FILETIME
        Public nFileSizeHigh As UInteger
        'changed all to uint, otherwise you run into unexpected overflow
        Public nFileSizeLow As UInteger
        '|
        Public dwReserved0 As UInteger
        '|
        Public dwReserved1 As UInteger
        'v
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> _
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=MAX_ALTERNATE)> _
        Public cAlternate As String
    End Structure

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function DeleteFile(lpFileName As String) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function DeleteFileA(<MarshalAs(UnmanagedType.LPStr)> lpFileName As String) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Function DeleteFileW(<MarshalAs(UnmanagedType.LPWStr)> lpFileName As String) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32", CharSet:=CharSet.Unicode)> _
    Public Function FindFirstFile(lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As IntPtr
    End Function

    <DllImport("kernel32", CharSet:=CharSet.Unicode)> _
    Public Function FindNextFile(hFindFile As IntPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Boolean
    End Function

    <DllImport("kernel32.dll")> _
    Public Function FindClose(hFindFile As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll")> _
    Private Function RemoveDirectory(lpPathName As String) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function



    Sub Main()

        Dim _arguments As String() = Environment.GetCommandLineArgs()

        If _arguments(1).ToLower = "-path" Then
            Console.WriteLine("Removing folders and files from folder " + _arguments(2))
            Call RecursivelyDeleteDirectory("\\?\" + _arguments(2))
        End If


        Console.ReadKey()
    End Sub


    Private Function RecursivelyDeleteDirectory(directoryPath As String) As Boolean
        Dim isDeleted As Boolean = False
        Dim isFailed As Boolean = False
        Dim INVALID_HANDLE_VALUE As New IntPtr(-1)

        Dim findData As WIN32_FIND_DATA
        Dim backslash As Char() = New Char() {"\"c}
        Dim findHandle As IntPtr
        Dim path As String = String.Empty

        findHandle = FindFirstFile(directoryPath.TrimEnd(backslash) & "\*", findData)
        If findHandle <> INVALID_HANDLE_VALUE Then

            Do
                path = String.Format("{0}\{1}", directoryPath.TrimEnd(backslash), findData.cFileName)

                If (findData.dwFileAttributes And FileAttributes.Directory) <> 0 Then
                    ' this is a directory
                    If findData.cFileName <> "." AndAlso findData.cFileName <> ".." Then

                        If Not RecursivelyDeleteDirectory(path) Then
                            ' we failed to delete a directory
                            isFailed = True
                            Exit Do
                        End If
                    End If
                Else
                    ' delete this file
                    If Not DeleteFileW(path) Then
                        isFailed = True
                        Exit Do
                    Else
                        Console.WriteLine("Deleting file " + path)
                    End If
                End If
            Loop While FindNextFile(findHandle, findData)

            FindClose(findHandle)

            If Not isFailed Then
                ' the directory should be empty now, delete directory

                isDeleted = RemoveDirectory(directoryPath)
                Console.WriteLine("Deleting folder " + directoryPath)
            End If
        End If

        Return isDeleted
    End Function
End Module
