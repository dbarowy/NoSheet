module Graph
    open System.Collections.Generic
    open SpreadsheetAST
    type DirectedAcyclicGraph(formulas: Dictionary<Address,Expression>) =
        // convert inputs into immutable maps, for thread-safety
        let fs = Seq.map (fun (pair: KeyValuePair<Address,Expression>) -> (pair.Key, pair.Value)) formulas |> Map.ofSeq

        // the set of all formulas
        let outputs = Map.toSeq fs |> Seq.map (fun (addr,_) -> addr) |> Set.ofSeq

        // a map of the input ranges for a formula output (i.e., Addr -> Set<Range>)
        let formula_ranges =
            Map.map (fun addr expr ->
                SpreadsheetUtility.GetRangesFromExpr(expr)
            ) fs

        // a map from a formula address to its (immediate, non-transitive) input dependencies
        // note that some input addresses represent data and others represent formulas
        let formula_input_map =
            Map.map (fun addr expr ->
                let ranges = formula_ranges.[addr]
                let addrs = SpreadsheetUtility.GetAddressesFromExpr(expr)
                let raddrs = Set.unionMany (Set.map (fun (r: Range) -> r.GetAddresses()) ranges)
                Set.union addrs raddrs
            ) fs

        // the set of all inputs
        let inputs = Map.toSeq formula_input_map |> Seq.map (fun (faddr, iaddrs) -> iaddrs) |> Set.unionMany

        // a map from inputs to formula outputs
        let input_formula_map =
            Set.map (fun iaddr ->
                let faddrs = Seq.filter (fun faddr ->
                                 Set.contains iaddr formula_input_map.[faddr]
                             ) (outputs) |> Set.ofSeq
                (iaddr, faddrs)
            ) inputs |> Set.toSeq |> Map.ofSeq

        // this returns addresses of all cells that provide input *data* for a formula
        // note that this computes the transitive closure of the "is input to" relation
        member self.GetInputDependencies(formula_address: Address) : Set<Address> =
            let rec GetInputs(f: Address) : Set<Address> =
                // if f is a formula then get the addresses
                // of its inputs
                if formula_input_map.ContainsKey f then
                    Set.map (fun input ->
                        GetInputs(input)
                    ) (formula_input_map.[f]) |> Set.unionMany
                // if f is not a formula then it IS an input
                else
                    set [f]
            GetInputs formula_address

        // this returns all output formulas that depend on a particular input
        // note that this computes the transitive closure of the "is output for" relation
        member self.GetOutputDependencies(input_address: Address) : Set<Address> =
            let rec GetOutputs(i: Address) : Set<Address> =
                // get all of the outputs for this input
                match input_formula_map.TryFind i with
                | Some(faddrs) ->
                    // union all of the outputs
                    Set.map (fun faddr -> GetOutputs faddr) faddrs |> Set.unionMany
                | None -> Set.empty

            GetOutputs input_address

        // this method returns an array of all of addresses of formulas.
        // if only_terminals = true, then we only return those formulas
        // that are not themselves inputs to other formulas
        member self.FormulaAddresses(only_terminals: bool) : Address[] =
            if only_terminals then
                Set.filter (fun addr ->
                    not (inputs.Contains addr)
                ) outputs |> Set.toArray
            else
                outputs |> Set.toArray

        // returns true if the element is a cell containing data
        member self.isData(addr: Address) : bool =
            inputs.Contains addr && not (outputs.Contains addr)

        // returns true if the element is a formula
        member self.isFormula(addr: Address) : bool = not (self.isData addr)

        // this method returns an array of homogenous input vectors
        // (a set of input addresses) for the computation
        member self.HomogeneousInputs : Range[] =
            // for now, we do what the old CheckCell did: just
            // return input ranges that include at least one
            // data cell
            Set.map (fun (addr: Address) ->
                // eliminate ranges that don't have at
                // least one data element
                Set.filter (fun (rng: Range) ->
                    Set.exists (fun elem ->
                        self.isData(elem)
                    ) (rng.GetAddresses())
                ) (formula_ranges.[addr])
            ) outputs |> Set.unionMany |> Set.toArray
