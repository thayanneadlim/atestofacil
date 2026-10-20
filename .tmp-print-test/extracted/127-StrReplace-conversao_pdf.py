OLD:
            if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
            if ($null -ne $excel) {
                try { $excel.Quit() } catch {}
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
            }

NEW:
            if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
            if ($null -ne $excel -and $owned.Count -gt 0) {
                try { $excel.Quit() } catch {}
            }
            if ($null -ne $excel) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
            }