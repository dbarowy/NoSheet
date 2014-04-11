module Graph
    open System.Collections.Generic
    open SpreadsheetAST
    type DirectedAcyclicGraph(formulas: Dictionary<Address,Expression>, data: Dictionary<Address,string>) =
        // convert inputs into immutable maps, for thread-safety
        let fs = Seq.map (fun (pair: KeyValuePair<Address,Expression>) -> (pair.Key, pair.Value)) formulas |> Map.ofSeq
        let ds = Seq.map (fun (pair: KeyValuePair<Address,string>) -> (pair.Key, pair.Value)) data |> Map.ofSeq

        let formula_addresses = Map.toSeq fs |> Seq.map (fun (addr,_) -> addr) |> Set.ofSeq

        // these are all of the input addresses for a formula output
        // note that some inputs are data and others are formulas
        let formula_inputs =
            Map.map (fun addr expr ->
                let ranges = SpreadsheetUtility.GetRangesFromExpr(expr)
                let addrs = SpreadsheetUtility.GetAddressesFromExpr(expr)
                let raddrs = Set.unionMany (Set.map (fun (r: Range) -> r.GetAddresses()) ranges)
                Set.union addrs raddrs
            ) fs

        // these are all of the outputs that depend on a particular input
        let cell_outputs =
            Map.map (fun iaddr _ ->
                Set.filter (fun faddr ->
                    formula_inputs.[faddr].Contains iaddr
                ) formula_addresses
            ) ds

        // this returns addresses of all cells that provide input *data* for a formula
        // note that this computes the transitive closure of the "is input to" relation
        member self.GetInputDependencies(formula_address: Address) : Set<Address> =
            let rec GetInputs(f: Address) : Set<Address> =
                // if f is a formula then get the addresses
                // of its inputs
                if formula_inputs.ContainsKey f then
                    Set.map (fun input ->
                        GetInputs(input)
                    ) (formula_inputs.[f]) |> Set.unionMany
                // if f is not a formula then it IS an input
                else
                    set [f]
            GetInputs formula_address

        // this returns all output formulas that depend on a particular input
        // note that this computes the transitive closure of the "is output for" relation
        member self.GetOutputDependencies(input_address: Address) : Set<Address> =
            match cell_outputs.TryFind input_address with
            | Some(addrs) -> addrs
            | None -> Set.empty

        // this method returns a set of all of addresses of formulas
        // if only_terminals = true, then we only return those formulas
        // that are not themselves inputs to other formulas
        member self.FormulaAddresses(only_terminals: bool) : Set<Address> =
            if only_terminals then
                Set.filter (fun addr ->
                    not (formula_inputs.ContainsKey addr)
                ) formula_addresses
            else
                formula_addresses

        member self.HomogeneousInputs : Set<Set<Address>> =
            
            failwith "nope"