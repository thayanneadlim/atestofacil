OLD:
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

try {

NEW:
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Resume-Recovery $recuperacao
    }
}

try {