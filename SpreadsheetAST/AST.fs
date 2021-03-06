﻿namespace SpreadsheetAST
    open System
    open System.Diagnostics

    type IndirectAddressingNotSupportedException(expression: string) =
        inherit Exception(expression)

    [<Serializable>]
    type Address() =
        let mutable R: int = 0
        let mutable C: int = 0
        let mutable _wsn = None
        let mutable _wbn = None
        let mutable _path = None        // this should actually be the enclosing directory!
        static member FromR1C1(R: int, C: int, wsname: string, wbname: string, path: string) : Address =
            let addr = Address()
            addr.Row <- R
            addr.Col <- C
            addr.WorksheetName <- Some(wsname)
            addr.WorkbookName <- Some(wbname)
            addr.Path <- Some(path)
            addr
        static member NewFromR1C1(R: int, C: int, wsname: string option, wbname: string option, path: string option) : Address =
            let addr = Address()
            addr.Row <- R
            addr.Col <- C
            addr.WorksheetName <- wsname
            addr.WorkbookName <- wbname
            addr.Path <- path
            addr
        static member NewFromA1(row: int, col: string, wsname: string option, wbname: string option, path: string option) : Address =
            let addr = Address()
            addr.Row <- row
            addr.Col <- Address.CharColToInt(col)
            addr.WorksheetName <- wsname
            addr.WorkbookName <- wbname
            addr.Path <- path
            addr
        member self.A1Local() : string = Address.IntToColChars(self.X) + self.Y.ToString()
        member self.A1Path() : string =
            match _path with
            | Some(pth) -> pth
            | None -> failwith "Path string should never be unset."
        member self.A1Worksheet() : string =
            match _wsn with
            | Some(ws) -> ws
            | None -> failwith "Worksheet string should never be unset."
        member self.A1Workbook() : string =
            match _wbn with
            | Some(wb) -> wb
            | None -> failwith "Workbook string should never be unset."
        member self.A1FullyQualified() : string =
            "[" + IO.Path.Combine(self.A1Path(), self.A1Workbook()) + "]" + self.A1Worksheet() + "!" + self.A1Local()
        member self.R1C1 =
            let wsstr = match _wsn with | Some(ws) -> ws + "!" | None -> ""
            let wbstr = match _wbn with | Some(wb) -> "[" + wb + "]" | None -> ""
            let pstr = match _path with | Some(pth) -> pth | None -> ""
            pstr + wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.X: int = C
        member self.Y: int = R
        member self.Row
            with get() = R
            and set(value) = R <- value
        member self.Col
            with get() = C
            and set(value) = C <- value
        member self.Path
            with get() = _path
            and set(value) = _path <- value
        member self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        member self.WorkbookName
            with get() = _wbn
            and set(value) = _wbn <- value
        // a Cantor pairing function for X and Y coordinates
        member private self.CantorPair : int =
            // adjust x and y to zero-based indexing
            let x = self.X - 1
            let y = self.Y - 1
            // calc
            ((x + y)*(x + y + 1))/2 + y
        member self.AddressAsInt32() =
            // convert to zero-based indices
            // the modulus catches overflow; collisions are OK because our equality
            // operator does an exact check
            // underflow should throw an exception
            let col_idx = (C - 1) % 65536       // allow 16 bits for columns
            let row_idx = (R - 1) % 65536       // allow 16 bits for rows
            if (col_idx < 0 || row_idx < 0) then
                failwith (System.String.Format("Nonsensical row or column index: {0}", self.ToString()))
            row_idx + (col_idx <<< 16)
        // Address is used as a Dictionary key, and reference equality
        // does not suffice, therefore GetHashCode and Equals are provided
        override self.GetHashCode() : int = self.AddressAsInt32()
        override self.Equals(obj: obj) : bool =
            let addr = obj :?> Address
            self.SameAs addr
        member self.SameAs(addr: Address) : bool =
            self.X = addr.X &&
            self.Y = addr.Y &&
            self.WorksheetName = addr.WorksheetName &&
            self.WorkbookName = addr.WorkbookName
        member self.InsideRange(rng: Range) : bool =
            not (self.X < rng.getXLeft() ||
                 self.Y < rng.getYTop() ||
                 self.X > rng.getXRight() ||
                 self.Y > rng.getYBottom())
        member self.InsideAddr(addr: Address) : bool =
            self.X = addr.X && self.Y = addr.Y
        override self.ToString() = self.A1FullyQualified()
        static member CharColToInt(col: string) : int =
            let rec ccti(idx: int) : int =
                let ltr = (int col.[idx]) - 64
                let num = (int (Math.Pow(26.0, float (col.Length - idx - 1)))) * ltr
                if idx = 0 then
                    num
                else
                    num + ccti(idx - 1)
            ccti(col.Length - 1)
        static member FromString(addr: string, wsname: string option, wbname: string option, path: string option) : Address =
            let reg = System.Text.RegularExpressions.Regex("R(?<row>[0-9]+)C(?<column>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = System.Convert.ToInt32(m.Groups.["column"].Value)
            Address.NewFromR1C1(r, c, wsname, wbname, path)
        static member IntToColChars(dividend: int) : string =
            let mutable quot = dividend / 26
            let rem = dividend % 26
            if rem = 0 then
                quot <- quot - 1
            let ltr = if rem = 0 then
                        'Z'
                      else
                        char (64 + rem)
            if quot = 0 then
                ltr.ToString()
            else
                Address.IntToColChars(quot) + ltr.ToString()
        interface IComparable with
            member self.CompareTo(obj: Object) : int =
                let addr = obj :?> Address
                self.A1FullyQualified().CompareTo(addr.A1FullyQualified())

    and IndirectAddress(expr: string) =
        inherit Address()
        do
            raise(IndirectAddressingNotSupportedException(expr))
        
    and Range(topleft: Address, bottomright: Address) =
        let _tl = topleft
        let _br = bottomright
        override self.ToString() =
            let tlstr = topleft.ToString()
            let brstr = bottomright.ToString()
            tlstr + "," + brstr
        member self.getXLeft() : int = _tl.X
        member self.getXRight() : int = _br.X
        member self.getYTop() : int = _tl.Y
        member self.getYBottom() : int = _br.Y
        member self.TopLeftAddress() : Address = _tl
        member self.BottomRightAddress() : Address = _br
        member self.GetAddresses() : Set<Address> =
            seq {
                for x in [_tl.X .. _br.X] do
                    for y in [_tl.Y .. _br.Y] do
                        yield Address.NewFromR1C1(y, x, _tl.WorksheetName, _tl.WorkbookName, _tl.Path)
            } |> Set.ofSeq
        member self.A1FullyQualified() : string =
            let wsname = _tl.A1Worksheet()
            let wbname = _tl.A1Workbook()
            let wbpath = _tl.A1Path()
            "[" + IO.Path.Combine(wbpath, wbname) + "]" + wsname + "!" + _tl.A1Local() + ":" + _br.A1Local()
        member self.InsideRange(rng: Range) : bool =
            if rng.TopLeftAddress().A1Path() <> self.TopLeftAddress().A1Path() ||
               rng.TopLeftAddress().A1Workbook() <> self.TopLeftAddress().A1Workbook() ||
               rng.TopLeftAddress().A1Worksheet() <> self.TopLeftAddress().A1Worksheet() then
               false
            else
                not (self.getXLeft() < rng.getXLeft() ||
                     self.getYTop() < rng.getYTop() ||
                     self.getXRight() > rng.getXRight() ||
                     self.getYBottom() > rng.getYBottom())
        member self.InsideAddr(addr: Address) : bool =
            if addr.A1Path() <> self.TopLeftAddress().A1Path() ||
               addr.A1Workbook() <> self.TopLeftAddress().A1Workbook() ||
               addr.A1Worksheet() <> self.TopLeftAddress().A1Worksheet() then
               false
            else
                not (self.getXLeft() < addr.X ||
                     self.getYTop() < addr.Y ||
                     self.getXRight() > addr.X ||
                     self.getYBottom() > addr.Y)
        member self.ContainsAddress(addr: Address) : bool =
            if addr.A1Path() <> self.TopLeftAddress().A1Path() ||
               addr.A1Workbook() <> self.TopLeftAddress().A1Workbook() ||
               addr.A1Worksheet() <> self.TopLeftAddress().A1Worksheet() then
               false
            else
                self.getXLeft() <= addr.X &&
                self.getXRight() >= addr.X &&
                self.getYTop() <= addr.Y &&
                self.getYBottom() >= addr.Y
        member self.SetPathName(path: string option) : unit =
            _tl.Path <- path
            _br.Path <- path
        member self.SetWorksheetName(wsname: string option) : unit =
            _tl.WorksheetName <- wsname
            _br.WorksheetName <- wsname
        member self.GetWorksheetName() : string =
            match _tl.WorksheetName with
            | Some(wsname) -> wsname
            | None -> match _br.WorksheetName with
                      | Some(wsname) -> wsname
                      | None -> failwith "No worksheet name set for range."
        member self.SetWorkbookName(wbname: string option) : unit =
            _tl.WorkbookName <- wbname
            _br.WorkbookName <- wbname
        member self.GetWorkbookNames() : seq<string> =
            [_tl.WorkbookName; _br.WorkbookName] |> List.choose id |> List.toSeq
        member self.GetPathNames() : seq<string> =
            [_tl.Path; _br.Path] |> List.choose id |> List.toSeq
        member self.Width : int = self.getXRight() - self.getXLeft() + 1
        member self.Height : int = self.getYBottom() - self.getYTop() + 1
        override self.Equals(obj: obj) : bool =
            let r = obj :?> Range
            self.getXLeft() = r.getXLeft() &&
            self.getXRight() = r.getXRight() &&
            self.getYTop() = r.getYTop() &&
            self.getYBottom() = r.getYBottom()
        interface IComparable with
            member self.CompareTo(obj: Object) : int =
                let rng = obj :?> Range
                self.A1FullyQualified().CompareTo(rng.A1FullyQualified())

    type ReferenceType =
    | ReferenceAddress  = 0
    | ReferenceRange    = 1
    | ReferenceFunction = 2
    | ReferenceConstant = 3
    | ReferenceString   = 4
    | ReferenceNamed    = 5

    [<AbstractClass>]
    type Reference(path: string option, wbname: string option, wsname: string option) =
        let mutable _path: string option = path
        let mutable _wbn: string option = wbname
        let mutable _wsn: string option = wsname
        abstract member InsideRef: Reference -> bool
        abstract member Path: string option with get, set
        // path -> workbook name -> worksheet name
        abstract member Resolve: string -> string -> string -> unit
        abstract member WorkbookName: string option with get, set
        abstract member WorksheetName: string option with get, set
        abstract member Type: ReferenceType
        default self.Path
            with get() = _path
            and set(value) = _path <- value
        default self.WorkbookName
            with get() = _wbn
            and set(value) = _wbn <- value
        default self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        default self.InsideRef(ref: Reference) = false
        default self.Resolve(path: string)(wbname: string)(wsname: string) : unit =
            // we assume that missing workbook and worksheet
            // names mean that the address is local to the current
            // workbook and worksheet
            _path <- match self.Path with
                     | Some(pth) -> Some pth
                     | None -> Some path
            _wbn <- match self.WorkbookName with
                    | Some(wbn) -> Some wbn
                    | None -> Some wbname
            _wsn <- match self.WorksheetName with
                    | Some(wsn) -> Some wsn
                    | None -> Some wsname

    and ReferenceRange(path: string option, wbname: string option, wsname: string option, rng: Range) =
        inherit Reference(path, wbname, wsname)
        do
            // also set the Workbook and Worksheet for the Range object itself
            rng.SetWorkbookName(wbname)
            rng.SetWorksheetName(wsname)
        override self.Type = ReferenceType.ReferenceRange
        override self.ToString() =
            let pth = match self.Path with
                      | Some(pth) -> pth
                      | None -> ""
            let wbn = match self.WorkbookName with
                      | Some(wbn) -> wbn
                      | None -> ""
            let wsn = match self.WorksheetName with
                      | Some(wsn) -> wsn
                      | None -> ""
            "ReferenceRange(" + pth + ",[" + wbn + "]," + wsn + "," + rng.ToString() + ")"
        override self.InsideRef(ref: Reference) : bool =
            match ref with
            | :? ReferenceAddress as ar -> rng.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> rng.InsideRange(rr.Range)
            | _ -> failwith "Unknown Reference subclass."
        member self.Range = rng
        override self.Resolve(path: string)(wbname: string)(wsname: string) =
            // we assume that missing workbook and worksheet
            // names mean that the address is local to the current
            // workbook and worksheet
            self.Path <- match self.Path with
                         | Some(pth) ->
                            rng.SetPathName(Some pth)
                            Some pth
                         | None ->
                            rng.SetPathName(Some path)
                            Some path
            self.WorkbookName <- match self.WorkbookName with
                                 // If we know it, we also pass the wbname
                                 // down to ranges and addresses
                                 | Some(wbn) ->
                                      rng.SetWorkbookName(Some wbn)
                                      Some wbn
                                 | None ->
                                      rng.SetWorkbookName(Some wbname)
                                      Some wbname
            self.WorksheetName <- match self.WorksheetName with
                                  | Some(wsn) ->
                                      rng.SetWorksheetName(Some wsn)
                                      Some wsn
                                  | None ->
                                      rng.SetWorksheetName(Some wsname)
                                      Some wsname
        override self.Equals(obj: obj) : bool =
            let rr = obj :?> ReferenceRange
            self.Path = rr.Path &&
            self.WorkbookName = rr.WorkbookName &&
            self.WorksheetName = rr.WorksheetName &&
            self.Range = rr.Range

    and ReferenceAddress(path: string option, wbname: string option, wsname: string option, addr: Address) =
        inherit Reference(path, wbname, wsname)
        do
            // also set the Workbook and Worksheet for the Address object itself
            addr.WorkbookName <- wbname
            addr.WorksheetName <- wsname
        override self.Type = ReferenceType.ReferenceAddress
        override self.ToString() =
            let pth = match self.Path with
                      | Some(pth) -> pth
                      | None -> ""
            let wbn = match self.WorkbookName with
                      | Some(wbn) -> wbn
                      | None -> ""
            let wsn = match self.WorksheetName with
                      | Some(wsn) -> wsn
                      | None -> ""
            "ReferenceAddress(" + pth + ",[" + wbn + "]," + wsn + "," + addr.ToString() + ")"
        member self.Address = addr
        override self.InsideRef(ref: Reference) =
            match ref with
            | :? ReferenceAddress as ar -> addr.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> addr.InsideRange(rr.Range)
            | _ -> failwith "Invalid Reference subclass."
        override self.Resolve(path: string)(wbname: string)(wsname: string) =
            // always resolve the workbook name when it is missing
            // but only resolve the worksheet name when the
            // workbook name is not set
            self.Path <- match self.Path with
                         | Some(pth) ->
                            addr.Path <- Some pth
                            Some pth
                         | None ->
                            addr.Path <- Some path
                            Some path
            self.WorkbookName <- match self.WorkbookName with
                                 // If we know it, we also pass the wbname
                                 // down to ranges and addresses
                                 | Some(wbn) ->
                                      addr.WorkbookName <- Some wbn
                                      Some wbn
                                 | None ->
                                      addr.WorkbookName <- Some wbname
                                      Some wbname
            self.WorksheetName <- match self.WorksheetName with
                                  | Some(wsn) ->
                                      addr.WorksheetName <- Some wsn
                                      Some wsn
                                  | None ->
                                      addr.WorksheetName <- Some wsname
                                      Some wsname
        override self.Equals(obj: obj) : bool =
            let ra = obj :?> ReferenceAddress
            self.Path = ra.Path &&
            self.WorkbookName = ra.WorkbookName &&
            self.WorksheetName = ra.WorksheetName &&
            self.Address = ra.Address

    and ReferenceFunction(wsname: string option, fnname: string, arglist: Expression list) =
        inherit Reference(None, None, wsname)
        override self.Type = ReferenceType.ReferenceFunction
        member self.ArgumentList = arglist
        member self.FunctionName = fnname
        override self.ToString() =
            fnname + "[function](" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
        override self.Resolve(path: string)(wbname: string)(wsname: string) =
            // pass wb and ws information down to arguments
            // wb and ws names do not matter for functions
            for expr in arglist do
                expr.Resolve path wbname wsname
        override self.Equals(obj: obj) : bool =
            let rf = obj :?> ReferenceFunction
            self.Path = rf.Path &&
            self.WorkbookName = rf.WorkbookName &&
            self.WorksheetName = rf.WorksheetName &&
            self.FunctionName = rf.FunctionName
            // TODO: should also check ArgumentList here!

    and ReferenceConstant(wsname: string option, value: double) =
        inherit Reference(None, None, wsname)
        override self.Type = ReferenceType.ReferenceConstant
        member self.Value = value
        override self.ToString() = "Constant(" + value.ToString() + ")"
        override self.Equals(obj: obj) : bool =
            let rc = obj :?> ReferenceConstant
            self.Path = rc.Path &&
            self.WorkbookName = rc.WorkbookName &&
            self.WorksheetName = rc.WorksheetName &&
            self.Value = rc.Value

    and ReferenceString(wsname: string option, value: string) =
        inherit Reference(None, None, wsname)
        override self.Type = ReferenceType.ReferenceString
        member self.Value = value
        override self.ToString() = "String(" + value + ")"
        override self.Equals(obj: obj) : bool =
            let rs = obj :?> ReferenceString
            self.Path = rs.Path &&
            self.WorkbookName = rs.WorkbookName &&
            self.WorksheetName = rs.WorksheetName &&
            self.Value = rs.Value

    and ReferenceNamed(wsname: string option, varname: string) =
        inherit Reference(None, None, wsname)
        override self.Type = ReferenceType.ReferenceNamed
        member self.Name = varname
        override self.ToString() =
            match self.WorksheetName with
            | Some(wsn) -> "ReferenceName(" + wsn + ", " + varname + ")"
            | None -> "ReferenceName(None, " + varname + ")"
        override self.Equals(obj: obj) : bool =
            let rn = obj :?> ReferenceNamed
            self.Path = rn.Path &&
            self.WorkbookName = rn.WorkbookName &&
            self.WorksheetName = rn.WorksheetName &&
            self.Name = rn.Name

    // TODO: implement .Equals!
    and Expression =
    | ReferenceExpr of Reference
    | BinOpExpr of string * Expression * Expression
    | UnaryOpExpr of char * Expression
    | ParensExpr of Expression
        member self.Resolve(path: string)(wbname: string)(wsname: string) =
            match self with
            | ReferenceExpr(r) -> r.Resolve path wbname wsname
            | BinOpExpr(op,e1,e2) ->
                e1.Resolve path wbname wsname
                e2.Resolve path wbname wsname
            | UnaryOpExpr(op, e) ->
                e.Resolve path wbname wsname
            | ParensExpr(e) ->
                e.Resolve path wbname wsname
        override self.ToString() =
            match self with
            | ReferenceExpr(r) -> "ReferenceExpr(" + r.ToString() + ")"
            | BinOpExpr(op,e1,e2) -> "BinOpExpr(" + op.ToString() + "," + e1.ToString() + "," + e2.ToString() + ")"
            | UnaryOpExpr(op, e) -> "UnaryOpExpr(" + op.ToString() + "," + e.ToString() + ")"
            | ParensExpr(e) -> "ParensExpr(" + e.ToString() + ")"