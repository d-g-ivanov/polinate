/*
 * Javascript Diff Algorithm
 *  By John Resig (http://ejohn.org/)
 *  Modified by Chu Alan "sprite"
 *
 * Released under the MIT license.
 *
 * More Info:
 *  http://ejohn.org/projects/javascript-diff-algorithm/
 */

 // https://johnresig.com/projects/javascript-diff-algorithm/
const diff = (function() {
    const schema = {
        del: {'font': {'strike': true, 'size': 12,'color': {'argb': '00ed2939'},'name': 'Calibri','scheme': 'minor'},'text': ''},
        ins: {'font': {'underline': true, 'size': 12,'color': {'argb': '004cbb17'},'name': 'Calibri','scheme': 'minor'},'text': ''},
        regular: {'font': {'size': 12, 'color': {'theme': 1},'name': 'Calibri','scheme': 'minor'},'text': ''}
    };

    return diffString;

    // for my needs
    function diffString( o, n ) {
        let str = _diffString(o, n);
        str = str.split( '{{split_here}}' )
        
        return str.reduce( (final, segment) => {
            if (!segment) return final;

            let res;
            // if ins
            if (segment.startsWith('<ins>') && segment.endsWith('</ins>')) {
                segment = segment.replace('<ins>', '').replace('</ins>', '');
                res = copy(schema.ins);
                res.text = segment;
            }
            // if del
            else if (segment.startsWith('<del>') && segment.endsWith('</del>')) {
                segment = segment.replace('<del>', '').replace('</del>', '');
                res = copy(schema.del);
                res.text = segment;
            }
            // else it is a normal string
            else {
                res = copy( schema.regular );
                res.text = segment;
            }
            
            final.push(res);
            return final;
        }, [])
    }

    // changed
    function _diffString( o, n ) {
        o = o.replace(/\s+$/, '');
        n = n.replace(/\s+$/, '');
        
        // out.o - non-object words are DELETIONS - they do not appear in the new string
        // out.n - non-object words are INSERTIONS - they do not appear in the original string
        var out = diff(o == "" ? [] : o.split(/\s+/), n == "" ? [] : n.split(/\s+/) );

        var str = "";
      
        // take care to return the same spacing as the original strings
        // end in new line character
        var oSpace = o.match(/\s+/g);
        if (oSpace == null) {
          oSpace = ["\n"];
        } else {
          oSpace.push("\n");
        }
        var nSpace = n.match(/\s+/g);
        if (nSpace == null) {
          nSpace = ["\n"];
        } else {
          nSpace.push("\n");
        }
                
        // if NEW is empty, wrap whole string in a deletion tags
        if (out.n.length == 0) {
            str += '{{split_here}}<del>';
            for (var i = 0; i < out.o.length; i++) {
              str += escape(out.o[i]) + oSpace[i];
            }
            str += '</del>{{split_here}}';
        }
        // if OLD is empty, wrap whole string in insertions tags?
        else if (out.o.length == 0) {
            str += '{{split_here}}<ins>';
            for (var i = 0; i < out.n.length; i++) {
              str += escape(out.n[i]) + nSpace[i];
            }
            str += '</ins>{{split_here}}';
        }
        // else there are both insertions and deletions
        else {
            // if NEW starts with a deletion
            if (out.n[0].text == null) {
                str += '{{split_here}}<del>';
                for (n = 0; n < out.o.length && out.o[n].text == null; n++) {
                    str += escape(out.o[n]) + oSpace[n];
                }
                str += '</del>{{split_here}}';
            }
        
            for ( var i = 0; i < out.n.length; i++ ) {
                if (out.n[i].text == null) {
                    str += '{{split_here}}<ins>' + escape(out.n[i]) + nSpace[i] + "</ins>{{split_here}}";
                } else {
                    var pre = "";
                    pre += '{{split_here}}<del>';
                    for (n = out.n[i].row + 1; n < out.o.length && out.o[n].text == null; n++ ) {
                        pre += escape(out.o[n]) + oSpace[n];
                    }
                    pre += '</del>{{split_here}}';

                    if (pre == '{{split_here}}<del></del>{{split_here}}') pre = "";

                    str += out.n[i].text + nSpace[i] + pre;
                    // str += " " + out.n[i].text + nSpace[i] + pre;
                }
            }
        }
        
        return str;
    }
    
    // original
    function __diffString( o, n ) {
      o = o.replace(/\s+$/, '');
      n = n.replace(/\s+$/, '');
    
      var out = diff(o == "" ? [] : o.split(/\s+/), n == "" ? [] : n.split(/\s+/) );
      var str = "";
    
      var oSpace = o.match(/\s+/g);
      if (oSpace == null) {
        oSpace = ["\n"];
      } else {
        oSpace.push("\n");
      }
      var nSpace = n.match(/\s+/g);
      if (nSpace == null) {
        nSpace = ["\n"];
      } else {
        nSpace.push("\n");
      }
    
      if (out.n.length == 0) {
          for (var i = 0; i < out.o.length; i++) {
            str += '<del>' + escape(out.o[i]) + oSpace[i] + "</del>";
          }
      } else {
        if (out.n[0].text == null) {
          for (n = 0; n < out.o.length && out.o[n].text == null; n++) {
            str += '<del>' + escape(out.o[n]) + oSpace[n] + "</del>";
          }
        }
    
        for ( var i = 0; i < out.n.length; i++ ) {
          if (out.n[i].text == null) {
            str += '<ins>' + escape(out.n[i]) + nSpace[i] + "</ins>";
          } else {
            var pre = "";
    
            for (n = out.n[i].row + 1; n < out.o.length && out.o[n].text == null; n++ ) {
              pre += '<del>' + escape(out.o[n]) + oSpace[n] + "</del>";
            }
            str += " " + out.n[i].text + nSpace[i] + pre;
          }
        }
      }
      
      return str;
    }
    
    function diff( o, n ) {
        var ns = new Object();
        var os = new Object();
        
        /*
            create maps for the words in each string
            - row flags the indices of the word in the string
        */
        // n = NEW, array made from a string, splitting by 1 or more spaces
        for ( var i = 0; i < n.length; i++ ) {
            if ( ns[ n[i] ] == null )
                ns[ n[i] ] = { rows: new Array(), o: null };
            ns[ n[i] ].rows.push( i );
        }
        // o = ORIGINAL, array made from a string, splitting by 1 or more spaces
        for ( var i = 0; i < o.length; i++ ) {
            if ( os[ o[i] ] == null )
                os[ o[i] ] = { rows: new Array(), n: null };
            os[ o[i] ].rows.push( i );
        }
        
        // GOAL: replace all UNIQUE words that are in both string
        /*
            use NEW map (ns) to:
            replace all words in n and o, where
            - the words exist in both n and 0, AND
            - the words occur only 1 time in both n and o
            with { text: WORD, row: INDEX }
        */
        for ( var i in ns ) {
            if ( ns[i].rows.length == 1 && typeof(os[i]) != "undefined" && os[i].rows.length == 1 ) {
                n[ ns[i].rows[0] ] = { text: n[ ns[i].rows[0] ], row: os[i].rows[0] };
                o[ os[i].rows[0] ] = { text: o[ os[i].rows[0] ], row: ns[i].rows[0] };
            }
        }

        // GOAL: eliminate duplicates?
        /*
            replace all words in n and o, where
            - current word has text prop
            - next word has no text prop
            - current word position should not exceed ORIGINAL length
            - ORIGINAL at index current word + 1 does not have text prop
            - next word is the same as the ORIGINAL word at the same index + 1
        */
        for ( var i = 0; i < n.length - 1; i++ ) {
            let currWord = n[i],
                nextWord = n[i+1];
            
            if (
                currWord.text != null               && nextWord.text == null                &&
                currWord.row + 1 < o.length         && o[ currWord.row + 1 ].text == null   &&
                nextWord == o[ currWord.row + 1 ]
            ) {
                n[i+1] = { text: n[i+1], row: n[i].row + 1 };
                o[n[i].row+1] = { text: o[n[i].row+1], row: i + 1 };
            }
        }
        
        /*
            same as above but in reverse order
        */
        for ( var i = n.length - 1; i > 0; i-- ) {
            let currWord = n[i],
                prevWord = n[i-1];
            
            if (
                currWord.text != null               && prevWord.text == null                &&
                currWord.row > 0                    && o[ currWord.row - 1 ]. text == null  &&
                prevWord == o[ currWord.row - 1 ]
            ) {
                n[i-1] = { text: n[i-1], row: n[i].row - 1 };
                o[n[i].row-1] = { text: o[n[i].row-1], row: i - 1 };
            }
        }
        
        return { o: o, n: n };
    }

    function escape(s) {
        var n = s;
        n = n.replace(/&/g, "&amp;");
        n = n.replace(/</g, "&lt;");
        n = n.replace(/>/g, "&gt;");
        n = n.replace(/"/g, "&quot;");
    
        return n;
    }

    function copy(obj) {
        return JSON.parse( JSON.stringify(obj) );
    }
})();


