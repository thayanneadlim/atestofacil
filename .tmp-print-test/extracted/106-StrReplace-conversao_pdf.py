OLD:
        } catch {
            $ultima = $_.Exception.Message
            if (Test-Path -LiteralPath $Destino) {

NEW:
        } catch {
            $ultima = $_.Exception.Message
            Add-Content -LiteralPath $Log -Value ("tentativa " + $tentativa + " " + $Origem + " " + $ultima) -Encoding utf8
            if (Test-Path -LiteralPath $Destino) {