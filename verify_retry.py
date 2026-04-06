import sys
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError
import winscript.core.retry_guard as rg1
import winscript.core.retry_guard as rg2

def test_retry_guard():
    tool = "test_tool"
    args = {"param": "value"}
    
    print("Testing record_failure...")
    for i in range(4):
        guard.record_failure(tool, args)
        print(f"Failed {i+1} times, no exception.")
        
    try:
        guard.record_failure(tool, args)
        print("ERROR: Expected WinScriptMaxRetriesError on 5th failure!")
        sys.exit(1)
    except WinScriptMaxRetriesError as e:
        print(f"Success: Caught expected exception on 5th failure: {e}")
        
    print("Testing counter reset...")
    guard.record_success(tool, args)
    
    for i in range(4):
        guard.record_failure(tool, args)
        print(f"Failed {i+1} times after reset, no exception.")
        
    print("Testing singleton identity...")
    if id(rg1.guard) == id(rg2.guard) == id(guard):
        print("Success: guard is a singleton across imports.")
    else:
        print("ERROR: guard is NOT a singleton.")
        sys.exit(1)

    print("All tests passed.")

if __name__ == "__main__":
    test_retry_guard()