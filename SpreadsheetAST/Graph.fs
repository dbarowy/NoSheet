module Graph
    open System.Collections.Generic
    open SpreadsheetAST
    type DirectedAcyclicGraph(formulas: Dictionary<Address,Expression>, data: Dictionary<Address,string>) =
        // these are all of the input addresses for a formula output
        // note that some inputs are data and others are formulas
        let formula_inputs =
            Seq.map (fun (pair: KeyValuePair<Address,Expression>) ->
                let addr = pair.Key
                let expr = pair.Value
                let ranges = SpreadsheetUtility.GetRangesFromExpr(expr)
                let addrs = SpreadsheetUtility.GetAddressesFromExpr(expr)
                let raddrs = Set.unionMany (Set.map (fun (r: Range) -> r.GetAddresses()) ranges)
                (addr, (Set.union addrs raddrs))
            ) formulas |> Map.ofSeq

        // these are all of the outputs that depend on a particular input
        let cell_outputs =
            Seq.map (fun (pair: KeyValuePair<Address,string>) ->
                let iaddr = pair.Key
                let outputs = Seq.map (fun (pair: KeyValuePair<Address,Expression>) ->
                                let faddr = pair.Key
                                if formula_inputs.[faddr].Contains(iaddr) then
                                    Some(faddr)
                                else
                                    None
                              ) formulas |> Seq.choose id |> Set.ofSeq
                (iaddr, outputs)
            ) data |> Map.ofSeq

        // this returns addresses of all cells that provide input *data* for a formula
        // note that this computes the transitive closure of the "is input to" relation
        member self.GetInputDependencies(formula_address: Address) : Set<Address> =
            let rec GetInputs(f: Address) : Set<Address> =
                // if f is a formula then get the addresses
                // of its inputs
                if formula_inputs.ContainsKey(f) then
                    Set.map (fun input ->
                        GetInputs(input)
                    ) (formula_inputs.[f]) |> Set.unionMany
                // if f is not a formula then it IS an input
                else
                    set [f]
            GetInputs(formula_address)

        // this returns all output formulas that depend on a particular input
        // note that this computes the transitive closure of the "is output for" relation
        member self.GetOutputDependencies(input_address: Address) : Set<Address> =
            match cell_outputs.TryFind input_address with
            | Some(addrs) -> addrs
            | None -> Set.empty

        // this returns only the ranges that are referenced directly in this formula
        // note: used for classic DataDebug algorithm
        member self.GetInputRanges(expr: Expression) : Set<Range> =
            SpreadsheetUtility.GetRangesFromExpr(expr)