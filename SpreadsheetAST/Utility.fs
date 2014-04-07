module SpreadsheetUtility
    open SpreadsheetAST
    open System.Collections.Generic

    let PuntedFunction(fnname: string) : bool =
        match fnname with
        | "INDEX" -> true
        | "HLOOKUP" -> true
        | "VLOOKUP" -> true
        | "LOOKUP" -> true
        | "OFFSET" -> true
        | _ -> false

    let rec GetRangesFromRangeReference(ref: ReferenceRange) : Set<Range> = set [ref.Range]

    and GetRangesFromFunction(ref: ReferenceFunction) : Set<Range> =
        if PuntedFunction(ref.FunctionName) then
            Set.empty
        else
            List.map (fun arg -> GetRangesFromExpr(arg)) ref.ArgumentList |> Set.unionMany
        
    and GetRangesFromExpr(expr: Expression) : Set<Range> =
        match expr with
        | ReferenceExpr(r) -> GetRangesFromRef(r)
        | BinOpExpr(op, e1, e2) -> Set.union (GetRangesFromExpr(e1)) (GetRangesFromExpr(e2))
        | UnaryOpExpr(op, e) -> GetRangesFromExpr(e)
        | ParensExpr(e) -> GetRangesFromExpr(e)

    and GetRangesFromRef(ref: Reference) : Set<Range> =
        match ref with
        | :? ReferenceRange as r -> GetRangesFromRangeReference(r)
        | :? ReferenceAddress -> Set.empty
        | :? ReferenceNamed -> Set.empty   // TODO: symbol table lookup
        | :? ReferenceFunction as r -> GetRangesFromFunction(r)
        | :? ReferenceConstant -> Set.empty
        | :? ReferenceString -> Set.empty
        | _ -> failwith "Unknown reference type."

    let rec GetAddressesFromExpr(expr: Expression) : Set<Address> =
        match expr with
        | ReferenceExpr(r) -> GetAddressesFromRef(r)
        | BinOpExpr(op, e1, e2) -> Set.union (GetAddressesFromExpr(e1)) (GetAddressesFromExpr(e2))
        | UnaryOpExpr(op, e) -> GetAddressesFromExpr(e)
        | ParensExpr(e) -> GetAddressesFromExpr(e)

    and GetAddressesFromRef(ref: Reference) : Set<Address> =
        match ref with
        | :? ReferenceRange -> Set.empty
        | :? ReferenceAddress as r -> GetAddressesFromAddressRef(r)
        | :? ReferenceNamed -> Set.empty   // TODO: symbol table lookup
        | :? ReferenceFunction as r -> GetAddressesFromFunction(r)
        | :? ReferenceConstant -> Set.empty
        | :? ReferenceString -> Set.empty
        | _ -> failwith "Unknown reference type."

    and GetAddressesFromAddressRef(ref: ReferenceAddress) : Set<Address> = set [ref.Address]

    and GetAddressesFromFunction(ref: ReferenceFunction) : Set<Address> =
        if PuntedFunction(ref.FunctionName) then
            Set.empty
        else
            List.map (fun arg -> GetAddressesFromExpr(arg)) ref.ArgumentList |> Set.unionMany

    let rec GetFormulaNamesFromExpr(ast: Expression): Set<string> =
        match ast with
        | ReferenceExpr(r) -> GetFormulaNamesFromReference(r)
        | BinOpExpr(op, e1, e2) -> Set.union (GetFormulaNamesFromExpr(e1)) (GetFormulaNamesFromExpr(e2))
        | UnaryOpExpr(op, e) -> GetFormulaNamesFromExpr(e)
        | ParensExpr(e) -> GetFormulaNamesFromExpr(e)

    and GetFormulaNamesFromReference(ref: Reference): Set<string> =
        match ref with
        | :? ReferenceFunction as r -> set [r.FunctionName]
        | _ -> Set.empty