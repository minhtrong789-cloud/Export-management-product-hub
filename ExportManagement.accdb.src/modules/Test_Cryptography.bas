Attribute VB_Name = "Test_Cryptography"
'@TestModule
'@Folder "Tests"

Option Compare Database

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object
Private Mocks As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set Mocks = CreateObject("Rubberduck.MockProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    Set Mocks = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.

End Sub

'====================================================================
' UNIT TESTS : Cryptography Engine
'====================================================================

'@TestMethod("Cryptography")
Private Sub Test_SHA256_EmptyString_ReturnsStandardDigest()
    On Error GoTo TestFail

    ' Arrange: Standard RFC SHA-256 digest for an empty string
    Dim inputStr As String: inputStr = ""
    Dim expectedDigest As String
    expectedDigest = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"

    ' Act: Compute hash using Windows CNG bcrypt provider
    Dim actualDigest As String
    actualDigest = Extension_Cryptography.SHA256_CNG(inputStr)

    ' Assert: Output must strictly match standard digest
    Assert.AreEqual expectedDigest, actualDigest
    Exit Sub

TestFail:
    Assert.Fail "Test raised an unexpected error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Cryptography")
Private Sub Test_SHA256_Deterministic_ReturnsIdenticalDigest()
    On Error GoTo TestFail

    ' Arrange: Sample payload for determinism check
    Dim sampleData As String: sampleData = "ExportManagement_Staging_12345"

    ' Act: Execute hash twice consecutively
    Dim firstHash As String: firstHash = Extension_Cryptography.SHA256_CNG(sampleData)
    Dim secondHash As String: secondHash = Extension_Cryptography.SHA256_CNG(sampleData)

    ' Assert: Hash must be 64-char hex and deterministic across calls
    Assert.IsTrue Len(firstHash) = 64
    Assert.AreEqual firstHash, secondHash
    Exit Sub

TestFail:
    Assert.Fail "Test raised an unexpected error: #" & Err.Number & " - " & Err.Description
End Sub
