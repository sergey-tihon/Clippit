// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/*Author:Martin.Holzherr;Date:20080922;Context:"PEG Support for C#";Licence:CPOL
 * <<History>>
 *  20080922;V1.0 created
 *  20080929;UTF16BE;Added UTF16BE read support to <<FileLoader.LoadFile(out string src)>>
 * <</History>>
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace Clippit.Excel
{
    #region Input File Support
    public enum EncodingClass { unicode, utf8, binary, ascii };
    public enum UnicodeDetection { notApplicable, BOM, FirstCharIsAscii };
    public class FileLoader
    {
        public enum FileEncoding { none, ascii, binary, utf8, unicode, utf16be, utf16le, utf32le, utf32be, uniCodeBOM };
        public FileLoader(EncodingClass encodingClass, UnicodeDetection detection, string path)
        {
            Encoding = GetEncoding(encodingClass, detection, path);
            _path = path;
        }
        public bool IsBinaryFile()
        {
            return Encoding == FileEncoding.binary;
        }
        public bool LoadFile(out byte[] src)
        {
            src = null;
            if (!IsBinaryFile()) return false;
            using var brdr = new BinaryReader(File.Open(_path, FileMode.Open,FileAccess.Read));
            src = brdr.ReadBytes((int)brdr.BaseStream.Length);
            return true;
        }
        public bool LoadFile(out string src)
        {
            src = null;
            var textEncoding = FileEncodingToTextEncoding();
            if (textEncoding == null)
            {
                if (Encoding == FileEncoding.binary) return false;
                using var rd = new StreamReader(_path, true);
                src = rd.ReadToEnd();
                return true;
            }
            else
            {
                if (Encoding == FileEncoding.utf16be)//UTF16BE
                {
                    using var brdr = new BinaryReader(File.Open(_path, FileMode.Open, FileAccess.Read));
                    var bytes = brdr.ReadBytes((int)brdr.BaseStream.Length);
                    var s = new StringBuilder();
                    for (var i = 0; i < bytes.Length; i += 2)
                    {
                        var c = (char)(bytes[i] << 8 | bytes[i + 1]);
                        s.Append(c);
                    }
                    src = s.ToString();
                    return true;
                }
                else
                {
                    using var rd = new StreamReader(_path, textEncoding);
                    src = rd.ReadToEnd();
                    return true;
                }
            }

        }

        private Encoding FileEncodingToTextEncoding()
        {
            switch (Encoding)
            {
                case FileEncoding.utf8: return new UTF8Encoding();
                case FileEncoding.utf32be:
                case FileEncoding.utf32le: return new UTF32Encoding();
                case FileEncoding.unicode:
                case FileEncoding.utf16be:
                case FileEncoding.utf16le: return new UnicodeEncoding();
                case FileEncoding.ascii: return new ASCIIEncoding();
                case FileEncoding.binary:
                case FileEncoding.uniCodeBOM: return null;
                default: Debug.Assert(false);
                    return null;

            }
        }

        private static FileEncoding DetermineUnicodeWhenFirstCharIsAscii(string path)
        {
            using var br = new BinaryReader(File.Open(path, FileMode.Open, FileAccess.Read));
            var startBytes = br.ReadBytes(4);
            if (startBytes.Length == 0) return FileEncoding.none;
            if (startBytes.Length is 1 or 3) return FileEncoding.utf8;
            if (startBytes.Length == 2 && startBytes[0] != 0) return FileEncoding.utf16le;
            if (startBytes.Length == 2 && startBytes[0] == 0) return FileEncoding.utf16be;
            if (startBytes[0] == 0 && startBytes[1] == 0) return FileEncoding.utf32be;
            if (startBytes[0] == 0 && startBytes[1] != 0) return FileEncoding.utf16be;
            if (startBytes[0] != 0 && startBytes[1] == 0) return FileEncoding.utf16le;
            return FileEncoding.utf8;
        }

        private FileEncoding GetEncoding(EncodingClass encodingClass, UnicodeDetection detection, string path)
        {
            return encodingClass switch
            {
                EncodingClass.ascii => FileEncoding.ascii,
                EncodingClass.unicode => detection switch
                {
                    UnicodeDetection.FirstCharIsAscii => DetermineUnicodeWhenFirstCharIsAscii(path),
                    UnicodeDetection.BOM => FileEncoding.uniCodeBOM,
                    _ => FileEncoding.unicode
                },
                EncodingClass.utf8 => FileEncoding.utf8,
                EncodingClass.binary => FileEncoding.binary,
                _ => FileEncoding.none
            };
        }

        private readonly string _path;
        private readonly FileEncoding Encoding;
    }
    #endregion Input File Support
    #region Error handling
    public class PegException : Exception
    {
        public PegException()
            : base("Fatal parsing error ocurred")
        {
        }
    }
    public struct PegError
    {
        internal SortedList<int, int> _lineStarts;

        private void AddLineStarts(string s, int first, int last, ref int lineNo, out int colNo)
        {
            colNo = 2;
            for (var i = first + 1; i <= last; ++i, ++colNo)
            {
                if (s[i - 1] == '\n')
                {
                    _lineStarts[i] = ++lineNo;
                    colNo = 1;
                }
            }
            --colNo;
        }
        public void GetLineAndCol(string s, int pos, out int lineNo, out int colNo)
        {
            for (var i = _lineStarts.Count; i > 0; --i)
            {
                var curLs = _lineStarts.ElementAt(i - 1);
                if (curLs.Key == pos)
                {
                    lineNo = curLs.Value;
                    colNo = 1;
                    return;
                }
                if (curLs.Key < pos)
                {
                    lineNo = curLs.Value;
                    AddLineStarts(s, curLs.Key, pos, ref lineNo, out colNo);
                    return;
                }
            }
            lineNo = 1;
            AddLineStarts(s, 0, pos, ref lineNo, out colNo);
        }
    }
    #endregion Error handling
    #region Syntax/Parse-Tree related classes
    public enum ESpecialNodes { eFatal = -10001, eAnonymNTNode = -1000, eAnonymASTNode = -1001, eAnonymousNode = -100 }
    public enum ECreatorPhase { eCreate, eCreationComplete, eCreateAndComplete }
    public struct PegBegEnd//indices into the source string
    {
        public int Length
        {
            get { return _posEnd - _posBeg; }
        }
        public string GetAsString(string src)
        {
            Debug.Assert(src.Length >= _posEnd);
            return src.Substring(_posBeg, Length);
        }
        public int _posBeg;
        public int _posEnd;
    }
    public class PegNode : ICloneable
    {
        #region Constructors
        public PegNode(PegNode parent, int id, PegBegEnd match, PegNode child, PegNode next)
        {
            parent_ = parent; id_ = id; child_ = child; next_ = next;
            match_ = match;
        }
        public PegNode(PegNode parent, int id, PegBegEnd match, PegNode child)
            : this(parent, id, match, child, null)
        {
        }
        public PegNode(PegNode parent, int id, PegBegEnd match)
            : this(parent, id, match, null, null)
        { }
        public PegNode(PegNode parent, int id)
            : this(parent, id, new PegBegEnd(), null, null)
        {
        }
        #endregion Constructors
        #region Public Members
        public virtual string GetAsString(string s)
        {
            return match_.GetAsString(s);
        }
        public virtual PegNode Clone()
        {
            var clone= new PegNode(parent_, id_, match_);
            CloneSubTrees(clone);
            return clone;
        }
        #endregion Public Members
        #region Protected Members
        protected void CloneSubTrees(PegNode clone)
        {
            PegNode child = null, next = null;
            if (child_ != null)
            {
                child = child_.Clone();
                child.parent_ = clone;
            }
            if (next_ != null)
            {
                next = next_.Clone();
                next.parent_ = clone;
            }
            clone.child_ = child;
            clone.next_ = next;
        }
        #endregion Protected Members
        #region Data Members
        public int id_;
        public PegNode parent_, child_, next_;
        public PegBegEnd match_;
        #endregion Data Members

        #region ICloneable Members

        object ICloneable.Clone()
        {
            return Clone();
        }

        #endregion
    }
    internal struct PegTree
    {
        internal enum AddPolicy { eAddAsChild, eAddAsSibling };
        internal PegNode _root;
        internal PegNode _cur;
        internal AddPolicy _addPolicy;
    }
    public abstract class PrintNode
    {
        public abstract int LenMaxLine();
        public abstract bool IsLeaf(PegNode p);
        public virtual bool IsSkip(PegNode p) { return false; }
        public abstract void PrintNodeBeg(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel);
        public abstract void PrintNodeEnd(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel);
        public abstract int LenNodeBeg(PegNode p);
        public abstract int LenNodeEnd(PegNode p);
        public abstract void PrintLeaf(PegNode p, ref int nOffsetLineBeg, bool bAlignVertical);
        public abstract int LenLeaf(PegNode p);
        public abstract int LenDistNext(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel);
        public abstract void PrintDistNext(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel);
    }
    public class TreePrint : PrintNode
    {
        #region Data Members
        public delegate string GetNodeName(PegNode node);

        private readonly string _src;
        private readonly TextWriter _treeOut;
        private readonly int _nMaxLineLen;
        private readonly bool _bVerbose;
        private readonly GetNodeName _getNodeName;
        #endregion Data Members
        #region Methods
        public TreePrint(TextWriter treeOut, string src, int nMaxLineLen, GetNodeName GetNodeName, bool bVerbose)
        {
            _treeOut = treeOut;
            _nMaxLineLen = nMaxLineLen;
            _bVerbose = bVerbose;
            _getNodeName = GetNodeName;
            _src = src;
        }

        public void PrintTree(PegNode parent, int nOffsetLineBeg, int nLevel)
        {
            if (IsLeaf(parent))
            {
                PrintLeaf(parent, ref nOffsetLineBeg, false);
                _treeOut.Flush();
                return;
            }
            var bAlignVertical =
                DetermineLineLength(parent, nOffsetLineBeg) > LenMaxLine();
            PrintNodeBeg(parent, bAlignVertical, ref nOffsetLineBeg, nLevel);
            var nOffset = nOffsetLineBeg;
            for (var p = parent.child_; p != null; p = p.next_)
            {
                if (IsSkip(p)) continue;

                if (IsLeaf(p))
                {
                    PrintLeaf(p, ref nOffsetLineBeg, bAlignVertical);
                }
                else
                {
                    PrintTree(p, nOffsetLineBeg, nLevel + 1);
                }
                if (bAlignVertical)
                {
                    nOffsetLineBeg = nOffset;
                }
                while (p.next_ != null && IsSkip(p.next_)) p = p.next_;

                if (p.next_ != null)
                {
                    PrintDistNext(p, bAlignVertical, ref nOffsetLineBeg, nLevel);
                }
            }
            PrintNodeEnd(parent, bAlignVertical, ref  nOffsetLineBeg, nLevel);
            _treeOut.Flush();
        }

        private int DetermineLineLength(PegNode parent, int nOffsetLineBeg)
        {
            var nLen = LenNodeBeg(parent);
            PegNode p;
            for (p = parent.child_; p != null; p = p.next_)
            {
                if (IsSkip(p)) continue;
                if (IsLeaf(p))
                {
                    nLen += LenLeaf(p);
                }
                else
                {
                    nLen += DetermineLineLength(p, nOffsetLineBeg);
                }
                if (nLen + nOffsetLineBeg > LenMaxLine())
                {
                    return nLen + nOffsetLineBeg;
                }
            }
            nLen += LenNodeEnd(p);
            return nLen;
        }
        public override int LenMaxLine() { return _nMaxLineLen; }
        public override void
            PrintNodeBeg(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel)
        {
            PrintIdAsName(p);
            _treeOut.Write("<");
            if (bAlignVertical)
            {
                _treeOut.WriteLine();
                _treeOut.Write(new string(' ', nOffsetLineBeg += 2));
            }
            else
            {
                ++nOffsetLineBeg;
            }
        }
        public override void
            PrintNodeEnd(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel)
        {
            if (bAlignVertical)
            {
                _treeOut.WriteLine();
                _treeOut.Write(new string(' ', nOffsetLineBeg -= 2));
            }
            _treeOut.Write('>');
            if (!bAlignVertical)
            {
                ++nOffsetLineBeg;
            }
        }
        public override int LenNodeBeg(PegNode p) { return LenIdAsName(p) + 1; }
        public override int LenNodeEnd(PegNode p) { return 1; }
        public override void PrintLeaf(PegNode p, ref int nOffsetLineBeg, bool bAlignVertical)
        {
            if (_bVerbose)
            {
                PrintIdAsName(p);
                _treeOut.Write('<');
            }
            var len = p.match_._posEnd - p.match_._posBeg;
            _treeOut.Write("'");
            if (len > 0)
            {
                _treeOut.Write(_src.Substring(p.match_._posBeg, p.match_._posEnd - p.match_._posBeg));
            }
            _treeOut.Write("'");
            if (_bVerbose) _treeOut.Write('>');
        }
        public override int LenLeaf(PegNode p)
        {
            var nLen = p.match_._posEnd - p.match_._posBeg + 2;
            if (_bVerbose) nLen += LenIdAsName(p) + 2;
            return nLen;
        }
        public override bool IsLeaf(PegNode p)
        {
            return p.child_ == null;
        }

        public override void
            PrintDistNext(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel)
        {
            if (bAlignVertical)
            {
                _treeOut.WriteLine();
                _treeOut.Write(new string(' ', nOffsetLineBeg));
            }
            else
            {
                _treeOut.Write(' ');
                ++nOffsetLineBeg;
            }
        }

        public override int
            LenDistNext(PegNode p, bool bAlignVertical, ref int nOffsetLineBeg, int nLevel)
        {
            return 1;
        }

        private int LenIdAsName(PegNode p) => _getNodeName(p).Length;

        private void PrintIdAsName(PegNode p) => _treeOut.Write(_getNodeName(p));

        #endregion Methods
    }
    #endregion Syntax/Parse-Tree related classes
    #region Parsers
    public abstract class PegBaseParser
    {
        #region Data Types
        public delegate bool Matcher();
        public delegate PegNode Creator(ECreatorPhase ePhase, PegNode parentOrCreated, int id);
        #endregion Data Types

        #region Data members
        protected int _srcLen;
        protected int _pos;
        private bool _bMute;
        protected TextWriter _errOut;
        protected Creator _nodeCreator;
        protected int _maxPos;
        private PegTree _tree;
        #endregion Data members

        public virtual string GetRuleNameFromId(int id)
        {
            //normally overridden
            return id switch
            {
                (int)ESpecialNodes.eFatal => "FATAL",
                (int)ESpecialNodes.eAnonymNTNode => "Nonterminal",
                (int)ESpecialNodes.eAnonymASTNode => "ASTNode",
                (int)ESpecialNodes.eAnonymousNode => "Node",
                _ => id.ToString()
            };
        }
        public virtual void GetProperties(out EncodingClass encoding, out UnicodeDetection detection)
        {
            encoding = EncodingClass.ascii;
            detection = UnicodeDetection.notApplicable;
        }
        public int GetMaximumPosition()
        {
            return _maxPos;
        }
        protected PegNode DefaultNodeCreator(ECreatorPhase phase, PegNode parentOrCreated, int id)
        {
            if (phase is ECreatorPhase.eCreate or ECreatorPhase.eCreateAndComplete)
                return new PegNode(parentOrCreated, id);

            if (parentOrCreated.match_._posEnd > _maxPos)
                _maxPos = parentOrCreated.match_._posEnd;
            return null;
        }
        #region Constructors
        public PegBaseParser(TextWriter errOut)
        {
            _srcLen = _pos = 0;
            _errOut = errOut;
            _nodeCreator = DefaultNodeCreator;
        }
        #endregion Constructors
        #region Reinitialization, TextWriter access,Tree Access

        public void Construct(TextWriter fOut)
        {
            _srcLen = _pos = 0;
            _bMute = false;
            SetErrorDestination(fOut);
            ResetTree();
        }

        public void Rewind() { _pos = 0; }

        public void SetErrorDestination(TextWriter errOut)
        {
            _errOut = errOut ?? new StreamWriter(Console.OpenStandardError());
        }

        #endregion Reinitialization, TextWriter access,Tree Access
         #region Tree root access, Tree Node generation/display
         public PegNode GetRoot() { return _tree._root; }
         public void ResetTree()
         {
             _tree._root = null;
             _tree._cur = null;
             _tree._addPolicy = PegTree.AddPolicy.eAddAsChild;
         }

         private void AddTreeNode(int nId, PegTree.AddPolicy newAddPolicy, Creator createNode, ECreatorPhase ePhase)
         {
             if (_bMute) return;
             if (_tree._root == null)
             {
                 _tree._root = _tree._cur = createNode(ePhase, _tree._cur, nId);
             }
             else if (_tree._addPolicy == PegTree.AddPolicy.eAddAsChild)
             {
                 _tree._cur = _tree._cur.child_ = createNode(ePhase, _tree._cur, nId);
             }
             else
             {
                 _tree._cur = _tree._cur.next_ = createNode(ePhase, _tree._cur.parent_, nId);
             }
             _tree._addPolicy = newAddPolicy;
         }

         private void RestoreTree(PegNode prevCur, PegTree.AddPolicy prevPolicy)
         {
             if (_bMute) return;
             if (prevCur == null)
             {
                 _tree._root = null;
             }
             else if (prevPolicy == PegTree.AddPolicy.eAddAsChild)
             {
                 prevCur.child_ = null;
             }
             else
             {
                 prevCur.next_ = null;
             }
             _tree._cur = prevCur;
             _tree._addPolicy = prevPolicy;
         }
         public bool TreeChars(Matcher toMatch)
         {
             return TreeCharsWithId((int)ESpecialNodes.eAnonymousNode, toMatch);
         }
         public bool TreeChars(Creator nodeCreator, Matcher toMatch)
         {
             return TreeCharsWithId(nodeCreator, (int)ESpecialNodes.eAnonymousNode, toMatch);
         }
         public bool TreeCharsWithId(int nId, Matcher toMatch)
         {
             return TreeCharsWithId(_nodeCreator, nId, toMatch);
         }
         public bool TreeCharsWithId(Creator nodeCreator, int nId, Matcher toMatch)
         {
             var pos = _pos;
             if (toMatch())
             {
                 if (!_bMute)
                 {
                     AddTreeNode(nId, PegTree.AddPolicy.eAddAsSibling, nodeCreator, ECreatorPhase.eCreateAndComplete);
                     _tree._cur.match_._posBeg = pos;
                     _tree._cur.match_._posEnd = _pos;
                 }
                 return true;
             }
             return false;
         }
         public bool TreeNT(int nRuleId, Matcher toMatch)
         {
             return TreeNT(_nodeCreator, nRuleId, toMatch);
         }
         public bool TreeNT(Creator nodeCreator, int nRuleId, Matcher toMatch)
         {
             if (_bMute) return toMatch();
             PegNode prevCur = _tree._cur, ruleNode;
             var prevPolicy = _tree._addPolicy;
             var posBeg = _pos;
             AddTreeNode(nRuleId, PegTree.AddPolicy.eAddAsChild, nodeCreator, ECreatorPhase.eCreate);
             ruleNode = _tree._cur;
             var bMatches = toMatch();
             if (!bMatches) RestoreTree(prevCur, prevPolicy);
             else
             {
                 ruleNode.match_._posBeg = posBeg;
                 ruleNode.match_._posEnd = _pos;
                 _tree._cur = ruleNode;
                 _tree._addPolicy = PegTree.AddPolicy.eAddAsSibling;
                 nodeCreator(ECreatorPhase.eCreationComplete, ruleNode, nRuleId);
             }
             return bMatches;
         }
         public bool TreeAST(int nRuleId, Matcher toMatch)
         {
             return TreeAST(_nodeCreator, nRuleId, toMatch);
         }
         public bool TreeAST(Creator nodeCreator, int nRuleId, Matcher toMatch)
         {
             if (_bMute) return toMatch();
             var bMatches = TreeNT(nodeCreator, nRuleId, toMatch);
             if (bMatches)
             {
                 if (_tree._cur.child_ != null && _tree._cur.child_.next_ == null && _tree._cur.parent_ != null)
                 {
                     if (_tree._cur.parent_.child_ == _tree._cur)
                     {
                         _tree._cur.parent_.child_ = _tree._cur.child_;
                         _tree._cur.child_.parent_ = _tree._cur.parent_;
                         _tree._cur = _tree._cur.child_;
                     }
                     else
                     {
                         PegNode prev;
                         for (prev = _tree._cur.parent_.child_; prev != null && prev.next_ != _tree._cur; prev = prev.next_)
                         {
                         }
                         if (prev != null)
                         {
                             prev.next_ = _tree._cur.child_;
                             _tree._cur.child_.parent_ = _tree._cur.parent_;
                             _tree._cur = _tree._cur.child_;
                         }
                     }
                 }
             }
             return bMatches;
         }
         public bool TreeNT(Matcher toMatch)
         {
             return TreeNT((int)ESpecialNodes.eAnonymNTNode, toMatch);
         }
         public bool TreeNT(Creator nodeCreator, Matcher toMatch)
         {
             return TreeNT(nodeCreator, (int)ESpecialNodes.eAnonymNTNode, toMatch);
         }
         public bool TreeAST(Matcher toMatch)
         {
             return TreeAST((int)ESpecialNodes.eAnonymASTNode, toMatch);
         }
         public bool TreeAST(Creator nodeCreator, Matcher toMatch)
         {
             return TreeAST(nodeCreator, (int)ESpecialNodes.eAnonymASTNode, toMatch);
         }
         public virtual string TreeNodeToString(PegNode node)
         {
             return GetRuleNameFromId(node.id_);
         }
         public void SetNodeCreator(Creator nodeCreator)
         {
             Debug.Assert(nodeCreator != null);
             _nodeCreator = nodeCreator;
         }
         #endregion Tree Node generation
         #region PEG  e1 e2 .. ; &e1 ; !e1 ;  e? ; e* ; e+ ; e{a,b} ; .
         public bool And(Matcher pegSequence)
         {
             var prevCur = _tree._cur;
             var prevPolicy = _tree._addPolicy;
             var pos0 = _pos;
             var bMatches = pegSequence();
             if (!bMatches)
             {
                 _pos = pos0;
                 RestoreTree(prevCur, prevPolicy);
             }
             return bMatches;
         }
         public bool Peek(Matcher toMatch)
         {
             var pos0 = _pos;
             var prevMute = _bMute;
             _bMute = true;
             var bMatches = toMatch();
             _bMute = prevMute;
             _pos = pos0;
             return bMatches;
         }
         public bool Not(Matcher toMatch)
         {
             var pos0 = _pos;
             var prevMute = _bMute;
             _bMute = true;
             var bMatches = toMatch();
             _bMute = prevMute;
             _pos = pos0;
             return !bMatches;
         }
         public bool PlusRepeat(Matcher toRepeat)
         {
             int i;
             for (i = 0; ; ++i)
             {
                 var pos0 = _pos;
                 if (!toRepeat())
                 {
                     _pos = pos0;
                     break;
                 }
             }
             return i > 0;
         }
         public bool OptRepeat(Matcher toRepeat)
         {
             for (; ; )
             {
                 var pos0 = _pos;
                 if (!toRepeat())
                 {
                     _pos = pos0;
                     return true;
                 }
             }
         }
         public bool Option(Matcher toMatch)
         {
             var pos0 = _pos;
             if (!toMatch()) _pos = pos0;
             return true;
         }
         public bool ForRepeat(int count, Matcher toRepeat)
         {
             var prevCur = _tree._cur;
             var prevPolicy = _tree._addPolicy;
             var pos0 = _pos;
             int i;
             for (i = 0; i < count; ++i)
             {
                 if (!toRepeat())
                 {
                     _pos = pos0;
                     RestoreTree(prevCur, prevPolicy);
                     return false;
                 }
             }
             return true;
         }
         public bool ForRepeat(int lower, int upper, Matcher toRepeat)
         {
             var prevCur = _tree._cur;
             var prevPolicy = _tree._addPolicy;
             var pos0 = _pos;
             int i;
             for (i = 0; i < upper; ++i)
             {
                 if (!toRepeat()) break;
             }
             if (i < lower)
             {
                 _pos = pos0;
                 RestoreTree(prevCur, prevPolicy);
                 return false;
             }
             return true;
         }
         public bool Any()
         {
             if (_pos < _srcLen)
             {
                 ++_pos;
                 return true;
             }
             return false;
         }
         #endregion PEG  e1 e2 .. ; &e1 ; !e1 ;  e? ; e* ; e+ ; e{a,b} ; .
    }
    public class PegByteParser : PegBaseParser
    {
        #region Data members
        private byte[] _src;
        private PegError _errors;
        #endregion Data members

        #region PEG optimizations
        public sealed class BytesetData
        {
            public struct Range
            {
                public Range(byte low, byte high) { this.low = low; this.high = high; }
                public byte low;
                public byte high;
            }

            private readonly System.Collections.BitArray charSet_;
            private readonly bool bNegated_;
            public BytesetData(System.Collections.BitArray b)
                : this(b, false)
            {
            }
            public BytesetData(System.Collections.BitArray b, bool bNegated)
            {
                charSet_ = new System.Collections.BitArray(b);
                bNegated_ = bNegated;
            }
            public BytesetData(Range[] r, byte[] c)
                : this(r, c, false)
            {
            }
            public BytesetData(Range[] r, byte[] c, bool bNegated)
            {
                var max = 0;
                if (r != null) foreach (var val in r) if (val.high > max) max = val.high;
                if (c != null) foreach (int val in c) if (val > max) max = val;
                charSet_ = new System.Collections.BitArray(max + 1, false);
                if (r != null)
                {
                    foreach (var val in r)
                    {
                        for (int i = val.low; i <= val.high; ++i)
                        {
                            charSet_[i] = true;
                        }
                    }
                }
                if (c != null) foreach (int val in c) charSet_[val] = true;
                bNegated_ = bNegated;
            }
            public bool Matches(byte c)
            {
                var bMatches = c < charSet_.Length && charSet_[(int)c];
                if (bNegated_) return !bMatches;
                else return bMatches;
            }
        }
        /*     public class BytesetData
             {
                 public struct Range
                 {
                     public Range(byte low, byte high) { this.low = low; this.high = high; }
                     public byte low;
                     public byte high;
                 }
                 protected System.Collections.BitArray charSet_;
                 bool bNegated_;
                 public BytesetData(System.Collections.BitArray b, bool bNegated)
                 {
                     charSet_ = new System.Collections.BitArray(b);
                     bNegated_ = bNegated;
                 }
                 public BytesetData(byte[] c, bool bNegated)
                 {
                     int max = 0;
                     foreach (int val in c) if (val > max) max = val;
                     charSet_ = new System.Collections.BitArray(max + 1, false);
                     foreach (int val in c) charSet_[val] = true;
                     bNegated_ = bNegated;
                 }
                 public BytesetData(Range[] r, byte[] c, bool bNegated)
                 {
                     int max = 0;
                     foreach (Range val in r) if (val.high > max) max = val.high;
                     foreach (int val in c) if (val > max) max = val;
                     charSet_ = new System.Collections.BitArray(max + 1, false);
                     foreach (Range val in r)
                     {
                         for (int i = val.low; i <= val.high; ++i)
                         {
                             charSet_[i] = true;
                         }
                     }
                     foreach (int val in c) charSet_[val] = true;
                 }


                 public bool Matches(byte c)
                 {
                     bool bMatches = c < charSet_.Length && charSet_[(int)c];
                     if (bNegated_) return !bMatches;
                     else return bMatches;
                 }
             }*/
        #endregion PEG optimizations
        #region Constructors
        public PegByteParser()
            : this(null)
        {
        }
        public PegByteParser(byte[] src):base(null)
        {
            SetSource(src);
        }
        public PegByteParser(byte[] src, TextWriter errOut):base(errOut)
        {
            SetSource(src);
        }
        #endregion Constructors
        #region Reinitialization, Source Code access, TextWriter access,Tree Access
        public void Construct(byte[] src, TextWriter fOut)
        {
            base.Construct(fOut);
            SetSource(src);
        }
        public void SetSource(byte[] src)
        {
            src ??= Array.Empty<byte>();
            _src = src; _srcLen = src.Length;
            _errors._lineStarts = new SortedList<int, int>();
            _errors._lineStarts[0] = 1;
        }
        public byte[] GetSource() { return _src; }

        #endregion Reinitialization, Source Code access, TextWriter access,Tree Access
        #region Setting host variables
        public bool Into(Matcher toMatch,out byte[] into)
        {
            var pos = _pos;
            if (toMatch())
            {
                var nLen = _pos - pos;
                into= new byte[nLen];
                for(var i=0;i<nLen;++i){
                    into[i] = _src[i+pos];
                }
                return true;
            }
            else
            {
                into = null;
                return false;
            }
        }
        public bool Into(Matcher toMatch,out PegBegEnd begEnd)
        {
            begEnd._posBeg = _pos;
            var bMatches = toMatch();
            begEnd._posEnd = _pos;
            return bMatches;
        }
        public bool Into(Matcher toMatch,out int into)
        {
            into = 0;
            if (!Into(toMatch,out byte[] s)) return false;
            into = 0;
            for (var i = 0; i < s.Length; ++i)
            {
                into <<= 8;
                into |= s[i];
            }
            return true;
        }
        public bool Into(Matcher toMatch,out double into)
        {
            into = 0.0;
            if (!Into(toMatch,out byte[] s)) return false;
            var encoding = System.Text.Encoding.UTF8;
            var sAsString = encoding.GetString(s);
            if (!double.TryParse(sAsString, out into)) return false;
            return true;
        }
        public bool BitsInto(int lowBitNo, int highBitNo,out int into)
        {
            if (_pos < _srcLen)
            {
                into = (_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1);
                ++_pos;
                return true;
            }
            into = 0;
            return false;
        }
        public bool BitsInto(int lowBitNo, int highBitNo, BytesetData toMatch, out int into)
        {
            if (_pos < _srcLen)
            {
                var value = (byte)((_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1));
                ++_pos;
                into = value;
                return toMatch.Matches(value);
            }
            into = 0;
            return false;
        }
        #endregion Setting host variables
        #region Error handling

        private void LogOutMsg(string sErrKind, string sMsg)
        {
            _errOut.WriteLine("<{0}>{1}:{2}", _pos, sErrKind, sMsg);
            _errOut.Flush();
        }
        public virtual bool Fatal(string sMsg)
        {

            LogOutMsg("FATAL", sMsg);
            throw new PegException();
        }
        public bool Warning(string sMsg)
        {
            LogOutMsg("WARNING", sMsg);
            return true;
        }
        #endregion Error handling
       #region PEG Bit level equivalents for PEG e1 ; &e1 ; !e1; e1:into ;
        public bool Bits(int lowBitNo, int highBitNo, byte toMatch)
        {
            if (_pos < _srcLen && ((_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1)) == toMatch)
            {
                ++_pos;
                return true;
            }
            return false;
        }
        public bool Bits(int lowBitNo, int highBitNo,BytesetData toMatch)
        {
            if( _pos < _srcLen )
            {
                var value= (byte)((_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1));
                ++_pos;
                return toMatch.Matches(value);
            }
            return false;
        }
        public bool PeekBits(int lowBitNo, int highBitNo, byte toMatch)
        {
            return _pos < _srcLen && ((_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1)) == toMatch;
        }
        public bool NotBits(int lowBitNo, int highBitNo, byte toMatch)
        {
            return !(_pos < _srcLen && ((_src[_pos] >> (lowBitNo - 1)) & ((1 << highBitNo) - 1)) == toMatch);
        }
        public bool IntoBits(int lowBitNo,int highBitNo,out int val)
        {
            return BitsInto(lowBitNo,highBitNo,out val);
        }
        public bool IntoBits(int lowBitNo, int highBitNo, BytesetData toMatch, out int val)
        {
            return BitsInto(lowBitNo, highBitNo, out val);
        }
        public bool Bit(int bitNo,byte toMatch)
        {
            if (_pos < _srcLen && ((_src[_pos]>>(bitNo-1))&1)==toMatch){
                ++_pos;
                return true;
            }
            return false;
        }
        public bool PeekBit(int bitNo, byte toMatch)
        {
            return _pos < _srcLen && ((_src[_pos] >> (bitNo - 1)) & 1) == toMatch;
        }
        public bool NotBit(int bitNo, byte toMatch)
        {
            return !(_pos < _srcLen && ((_src[_pos] >> (bitNo - 1)) & 1) == toMatch);
        }
        #endregion PEG Bit level equivalents for PEG e1 ; &e1 ; !e1; e1:into ;
        #region PEG '<Literal>' / '<Literal>'/i / [low1-high1,low2-high2..] / [<CharList>]
        public bool Char(byte c1)
        {
            if (_pos < _srcLen && _src[_pos] == c1)
            { ++_pos; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2)
        {
            if (_pos + 1 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2)
            { _pos += 2; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3)
        {
            if (_pos + 2 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3)
            { _pos += 3; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3, byte c4)
        {
            if (_pos + 3 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4)
            { _pos += 4; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3, byte c4, byte c5)
        {
            if (_pos + 4 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5)
            { _pos += 5; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3, byte c4, byte c5, byte c6)
        {
            if (_pos + 5 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6)
            { _pos += 6; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3, byte c4, byte c5, byte c6, byte c7)
        {
            if (_pos + 6 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6
                && _src[_pos + 6] == c7)
            { _pos += 7; return true; }
            return false;
        }
        public bool Char(byte c1, byte c2, byte c3, byte c4, byte c5, byte c6, byte c7, byte c8)
        {
            if (_pos + 7 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6
                && _src[_pos + 6] == c7
                && _src[_pos + 7] == c8)
            { _pos += 8; return true; }
            return false;
        }
        public bool Char(byte[] s)
        {
            var sLength = s.Length;
            if (_pos + sLength > _srcLen) return false;
            for (var i = 0; i < sLength; ++i)
            {
                if (s[i] != _src[_pos + i]) return false;
            }
            _pos += sLength;
            return true;
        }
        public static byte ToUpper(byte c)
        {
            if (c >= 97 && c <= 122) return (byte)(c - 32); else return c;
        }
        public bool IChar(byte c1)
        {
            if (_pos < _srcLen && ToUpper(_src[_pos]) == c1)
            { ++_pos; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2)
        {
            if (_pos + 1 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2))
            { _pos += 2; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2, byte c3)
        {
            if (_pos + 2 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2)
                && ToUpper(_src[_pos + 2]) == ToUpper(c3))
            { _pos += 3; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2, byte c3, byte c4)
        {
            if (_pos + 3 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2)
                && ToUpper(_src[_pos + 2]) == ToUpper(c3)
                && ToUpper(_src[_pos + 3]) == ToUpper(c4))
            { _pos += 4; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2, byte c3, byte c4, byte c5)
        {
            if (_pos + 4 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2)
                && ToUpper(_src[_pos + 2]) == ToUpper(c3)
                && ToUpper(_src[_pos + 3]) == ToUpper(c4)
                && ToUpper(_src[_pos + 4]) == ToUpper(c5))
            { _pos += 5; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2, byte c3, byte c4, byte c5, byte c6)
        {
            if (_pos + 5 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2)
                && ToUpper(_src[_pos + 2]) == ToUpper(c3)
                && ToUpper(_src[_pos + 3]) == ToUpper(c4)
                && ToUpper(_src[_pos + 4]) == ToUpper(c5)
                && ToUpper(_src[_pos + 5]) == ToUpper(c6))
            { _pos += 6; return true; }
            return false;
        }
        public bool IChar(byte c1, byte c2, byte c3, byte c4, byte c5, byte c6, byte c7)
        {
            if (_pos + 6 < _srcLen
                && ToUpper(_src[_pos]) == ToUpper(c1)
                && ToUpper(_src[_pos + 1]) == ToUpper(c2)
                && ToUpper(_src[_pos + 2]) == ToUpper(c3)
                && ToUpper(_src[_pos + 3]) == ToUpper(c4)
                && ToUpper(_src[_pos + 4]) == ToUpper(c5)
                && ToUpper(_src[_pos + 5]) == ToUpper(c6)
                && ToUpper(_src[_pos + 6]) == ToUpper(c7))
            { _pos += 7; return true; }
            return false;
        }
        public bool IChar(byte[] s)
        {
            var sLength = s.Length;
            if (_pos + sLength > _srcLen) return false;
            for (var i = 0; i < sLength; ++i)
            {
                if (s[i] != ToUpper(_src[_pos + i])) return false;
            }
            _pos += sLength;
            return true;
        }
        public bool In(byte c0, byte c1)
        {
            if (_pos < _srcLen
                && _src[_pos] >= c0 && _src[_pos] <= c1)
            {
                ++_pos;
                return true;
            }
            return false;
        }
        public bool In(byte c0, byte c1, byte c2, byte c3)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(byte c0, byte c1, byte c2, byte c3, byte c4, byte c5)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3
                    || c >= c4 && c <= c5)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(byte c0, byte c1, byte c2, byte c3, byte c4, byte c5, byte c6, byte c7)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3
                    || c >= c4 && c <= c5
                    || c >= c6 && c <= c7)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(byte[] s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length - 1; i += 2)
                {
                    if (c >= s[i] && c <= s[i + 1])
                    {
                        ++_pos;
                        return true;
                    }
                }
            }
            return false;
        }
        public bool NotIn(byte[] s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length - 1; i += 2)
                {
                    if ( c >= s[i] && c <= s[i + 1] ) return false;
                }
                ++_pos;
                return true;
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1)
        {
            if (_pos < _srcLen
                && (_src[_pos] == c0 || _src[_pos] == c1))
            {
                ++_pos;
                return true;
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2, byte c3)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2, byte c3, byte c4)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2, byte c3, byte c4, byte c5)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2, byte c3, byte c4, byte c5, byte c6)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5 || c == c6)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte c0, byte c1, byte c2, byte c3, byte c4, byte c5, byte c6, byte c7)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5 || c == c6 || c == c7)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(byte[] s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length; ++i)
                {
                    if (c == s[i]) { ++_pos; return true; }
                }
            }
            return false;
        }
        public bool NotOneOf(byte[] s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length; ++i)
                {
                    if (c == s[i]) { return false; }
                }
                return true;
            }
            return false;
        }
        public bool OneOf(BytesetData bset)
        {
            if(_pos < _srcLen && bset.Matches(_src[_pos]))
            {
                ++_pos; return true;
            }
            return false;
        }
        #endregion PEG '<Literal>' / '<Literal>'/i / [low1-high1,low2-high2..] / [<CharList>]
    }
    public class PegCharParser : PegBaseParser
    {
        #region Data members
        protected string _src;
        private PegError _errors;
        #endregion Data members
        #region PEG optimizations
        public sealed class OptimizedCharset
        {
            public struct Range
            {
                public Range(char low, char high) { this.low = low; this.high = high; }
                public char low;
                public char high;
            }

            private readonly System.Collections.BitArray charSet_;
            private readonly bool bNegated_;
            public OptimizedCharset(System.Collections.BitArray b)
                : this(b, false)
            {
            }
            public OptimizedCharset(System.Collections.BitArray b, bool bNegated)
            {
                charSet_ = new System.Collections.BitArray(b);
                bNegated_ = bNegated;
            }
            public OptimizedCharset(Range[] r, char[] c)
                : this(r, c, false)
            {
            }
            public OptimizedCharset(Range[] r, char[] c, bool bNegated)
            {
                var max = 0;
                if (r != null) foreach (var val in r) if (val.high > max) max = val.high;
                if (c != null) foreach (int val in c) if (val > max) max = val;
                charSet_ = new System.Collections.BitArray(max + 1, false);
                if (r != null)
                {
                    foreach (var val in r)
                    {
                        for (int i = val.low; i <= val.high; ++i)
                        {
                            charSet_[i] = true;
                        }
                    }
                }
                if (c != null) foreach (int val in c) charSet_[val] = true;
                bNegated_ = bNegated;
            }


            public bool Matches(char c)
            {
                var bMatches = c < charSet_.Length && charSet_[(int)c];
                if (bNegated_) return !bMatches;
                else return bMatches;
            }
        }
        public sealed class OptimizedLiterals
        {
            internal class Trie
            {
                internal Trie(char cThis,int nIndex, string[] literals)
                {
                    cThis_ = cThis;
                    var cMax = char.MinValue;
                    cMin_ = char.MaxValue;
                    var followChars = new HashSet<char>();

                    foreach (var literal in literals)
                    {
                        if (literal==null ||  nIndex > literal.Length ) continue;
                        if (nIndex == literal.Length)
                        {
                            bLitEnd_ = true;
                            continue;
                        }
                        var c = literal[nIndex];
                        followChars.Add(c);
                        if ( c < cMin_) cMin_ = c;
                        if ( c > cMax) cMax = c;
                    }
                    if (followChars.Count == 0)
                    {
                        children_ = null;
                    }
                    else
                    {
                        children_ = new Trie[(cMax - cMin_) + 1];
                        foreach (var c in followChars)
                        {
                            var subLiterals = new List<string>();
                            foreach (var s in literals)
                            {
                                if ( nIndex >= s.Length ) continue;
                                if (c == s[nIndex])
                                {
                                    subLiterals.Add(s);
                                }
                            }
                            children_[c - cMin_] = new Trie(c, nIndex + 1, subLiterals.ToArray());
                        }
                    }

                }
                internal char cThis_;           //character stored in this node
                internal bool bLitEnd_;         //end of literal

                internal char cMin_;            //first valid character in children
                internal Trie[] children_;      //contains the successor node of cThis_;
            }
            internal Trie literalsRoot;
            public OptimizedLiterals(string[] litAlternatives)
            {
                literalsRoot = new Trie('\u0000', 0, litAlternatives);
            }
        }
        #endregion  PEG optimizations
        #region Constructors
        public PegCharParser():this("")
        {


        }
        public PegCharParser(string src):base(null)
        {
            SetSource(src);
        }
        public PegCharParser(string src, TextWriter errOut):base(errOut)
        {
            SetSource(src);
            _nodeCreator = DefaultNodeCreator;
        }
        #endregion Constructors
        #region Overrides
        public override string TreeNodeToString(PegNode node)
        {
            var label = base.TreeNodeToString(node);
            if (node.id_ == (int)ESpecialNodes.eAnonymousNode)
            {
                var value = node.GetAsString(_src);
                if (value.Length < 32) label += " <" + value + ">";
                else label += " <" + value.Substring(0, 29) + "...>";
            }
            return label;
        }
        #endregion Overrides
        #region Reinitialization, Source Code access, TextWriter access,Tree Access
        public void Construct(string src, TextWriter Fout)
        {
            base.Construct(Fout);
            SetSource(src);
        }
        public void SetSource(string src)
        {
            if (src == null) src = "";
            _src = src; _srcLen = src.Length; _pos = 0;
            _errors._lineStarts = new SortedList<int, int>();
            _errors._lineStarts[0] = 1;
        }
        public string GetSource() { return _src; }
        #endregion Reinitialization, Source Code access, TextWriter access,Tree Access
        #region Setting host variables
        public bool Into(Matcher toMatch,out string into)
        {
            var pos = _pos;
            if (toMatch())
            {
                into = _src.Substring(pos, _pos - pos);
                return true;
            }
            else
            {
                into = "";
                return false;
            }
        }
        public bool Into(Matcher toMatch,out PegBegEnd begEnd)
        {
            begEnd._posBeg = _pos;
            var bMatches = toMatch();
            begEnd._posEnd = _pos;
            return bMatches;
        }
        public bool Into(Matcher toMatch,out int into)
        {
            into = 0;
            if (!Into(toMatch,out string s)) return false;
            if (!int.TryParse(s, out into)) return false;
            return true;
        }
        public bool Into(Matcher toMatch,out double into)
        {
            into = 0.0;
            if (!Into(toMatch,out string s)) return false;
            if (!double.TryParse(s, out into)) return false;
            return true;
        }
        #endregion Setting host variables
        #region Error handling

        private void LogOutMsg(string sErrKind, string sMsg)
        {
            _errors.GetLineAndCol(_src, _pos, out var lineNo, out var colNo);
            _errOut.WriteLine("<{0},{1},{2}>{3}:{4}", lineNo, colNo, _maxPos, sErrKind, sMsg);
            _errOut.Flush();
        }
        public virtual bool Fatal(string sMsg)
        {

            LogOutMsg("FATAL", sMsg);
            throw new PegException();
            //return false;
        }
        public bool Warning(string sMsg)
        {
            LogOutMsg("WARNING", sMsg);
            return true;
        }
        #endregion Error handling
        #region PEG  optimized version of  e* ; e+
        public bool OptRepeat(OptimizedCharset charset)
        {
            for (; _pos < _srcLen && charset.Matches(_src[_pos]); ++_pos) ;
            return true;
        }
        public bool PlusRepeat(OptimizedCharset charset)
        {
            var pos0 = _pos;
            for (; _pos < _srcLen && charset.Matches(_src[_pos]); ++_pos) ;
            return _pos > pos0;
        }
        #endregion PEG  optimized version of  e* ; e+
        #region PEG '<Literal>' / '<Literal>'/i / [low1-high1,low2-high2..] / [<CharList>]
        public bool Char(char c1)
        {
            if (_pos < _srcLen && _src[_pos] == c1)
            { ++_pos; return true; }
            return false;
        }
        public bool Char(char c1, char c2)
        {
            if (_pos + 1 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2)
            { _pos += 2; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3)
        {
            if (_pos + 2 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3)
            { _pos += 3; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3, char c4)
        {
            if (_pos + 3 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4)
            { _pos += 4; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3, char c4, char c5)
        {
            if (_pos + 4 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5)
            { _pos += 5; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3, char c4, char c5, char c6)
        {
            if (_pos + 5 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6)
            { _pos += 6; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3, char c4, char c5, char c6, char c7)
        {
            if (_pos + 6 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6
                && _src[_pos + 6] == c7)
            { _pos += 7; return true; }
            return false;
        }
        public bool Char(char c1, char c2, char c3, char c4, char c5, char c6, char c7, char c8)
        {
            if (_pos + 7 < _srcLen
                && _src[_pos] == c1
                && _src[_pos + 1] == c2
                && _src[_pos + 2] == c3
                && _src[_pos + 3] == c4
                && _src[_pos + 4] == c5
                && _src[_pos + 5] == c6
                && _src[_pos + 6] == c7
                && _src[_pos + 7] == c8)
            { _pos += 8; return true; }
            return false;
        }
        public bool Char(string s)
        {
            var sLength = s.Length;
            if (_pos + sLength > _srcLen) return false;
            for (var i = 0; i < sLength; ++i)
            {
                if (s[i] != _src[_pos + i]) return false;
            }
            _pos += sLength;
            return true;
        }
        public bool IChar(char c1)
        {
            if (_pos < _srcLen && char.ToUpper(_src[_pos]) == c1)
            { ++_pos; return true; }
            return false;
        }
        public bool IChar(char c1, char c2)
        {
            if (_pos + 1 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2))
            { _pos += 2; return true; }
            return false;
        }
        public bool IChar(char c1, char c2, char c3)
        {
            if (_pos + 2 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2)
                && char.ToUpper(_src[_pos + 2]) == char.ToUpper(c3))
            { _pos += 3; return true; }
            return false;
        }
        public bool IChar(char c1, char c2, char c3, char c4)
        {
            if (_pos + 3 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2)
                && char.ToUpper(_src[_pos + 2]) == char.ToUpper(c3)
                && char.ToUpper(_src[_pos + 3]) == char.ToUpper(c4))
            { _pos += 4; return true; }
            return false;
        }
        public bool IChar(char c1, char c2, char c3, char c4, char c5)
        {
            if (_pos + 4 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2)
                && char.ToUpper(_src[_pos + 2]) == char.ToUpper(c3)
                && char.ToUpper(_src[_pos + 3]) == char.ToUpper(c4)
                && char.ToUpper(_src[_pos + 4]) == char.ToUpper(c5))
            { _pos += 5; return true; }
            return false;
        }
        public bool IChar(char c1, char c2, char c3, char c4, char c5, char c6)
        {
            if (_pos + 5 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2)
                && char.ToUpper(_src[_pos + 2]) == char.ToUpper(c3)
                && char.ToUpper(_src[_pos + 3]) == char.ToUpper(c4)
                && char.ToUpper(_src[_pos + 4]) == char.ToUpper(c5)
                && char.ToUpper(_src[_pos + 5]) == char.ToUpper(c6))
            { _pos += 6; return true; }
            return false;
        }
        public bool IChar(char c1, char c2, char c3, char c4, char c5, char c6, char c7)
        {
            if (_pos + 6 < _srcLen
                && char.ToUpper(_src[_pos]) == char.ToUpper(c1)
                && char.ToUpper(_src[_pos + 1]) == char.ToUpper(c2)
                && char.ToUpper(_src[_pos + 2]) == char.ToUpper(c3)
                && char.ToUpper(_src[_pos + 3]) == char.ToUpper(c4)
                && char.ToUpper(_src[_pos + 4]) == char.ToUpper(c5)
                && char.ToUpper(_src[_pos + 5]) == char.ToUpper(c6)
                && char.ToUpper(_src[_pos + 6]) == char.ToUpper(c7))
            { _pos += 7; return true; }
            return false;
        }
        public bool IChar(string s)
        {
            var sLength = s.Length;
            if (_pos + sLength > _srcLen) return false;
            for (var i = 0; i < sLength; ++i)
            {
                if (s[i] != char.ToUpper(_src[_pos + i])) return false;
            }
            _pos += sLength;
            return true;
        }

        public bool In(char c0, char c1)
        {
            if (_pos < _srcLen
                && _src[_pos] >= c0 && _src[_pos] <= c1)
            {
                ++_pos;
                return true;
            }
            return false;
        }
        public bool In(char c0, char c1, char c2, char c3)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(char c0, char c1, char c2, char c3, char c4, char c5)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3
                    || c >= c4 && c <= c5)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(char c0, char c1, char c2, char c3, char c4, char c5, char c6, char c7)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c >= c0 && c <= c1
                    || c >= c2 && c <= c3
                    || c >= c4 && c <= c5
                    || c >= c6 && c <= c7)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool In(string s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length - 1; i += 2)
                {
                    if (!(c >= s[i] && c <= s[i + 1])) return false;
                }
                ++_pos;
                return true;
            }
            return false;
        }
        public bool NotIn(string s)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                for (var i = 0; i < s.Length - 1; i += 2)
                {
                    if ( c >= s[i] && c <= s[i + 1]) return false;
                }
                ++_pos;
                return true;
            }
            return false;
        }
        public bool OneOf(char c0, char c1)
        {
            if (_pos < _srcLen
                && (_src[_pos] == c0 || _src[_pos] == c1))
            {
                ++_pos;
                return true;
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2, char c3)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2, char c3, char c4)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2, char c3, char c4, char c5)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2, char c3, char c4, char c5, char c6)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5 || c == c6)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(char c0, char c1, char c2, char c3, char c4, char c5, char c6, char c7)
        {
            if (_pos < _srcLen)
            {
                var c = _src[_pos];
                if (c == c0 || c == c1 || c == c2 || c == c3 || c == c4 || c == c5 || c == c6 || c == c7)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(string s)
        {
            if (_pos < _srcLen)
            {
                if (s.IndexOf(_src[_pos]) != -1)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool NotOneOf(string s)
        {
            if (_pos < _srcLen)
            {
                if (s.IndexOf(_src[_pos]) == -1)
                {
                    ++_pos;
                    return true;
                }
            }
            return false;
        }
        public bool OneOf(OptimizedCharset cset)
        {
            if (_pos < _srcLen && cset.Matches(_src[_pos]))
            {
                ++_pos; return true;
            }
            return false;
        }
        public bool OneOfLiterals(OptimizedLiterals litAlt)
        {
            var node = litAlt.literalsRoot;
            var matchPos = _pos-1;
            for (var pos = _pos; pos < _srcLen ; ++pos)
            {
                var c = _src[pos];
                if (    node.children_==null
                    ||  c < node.cMin_ || c > node.cMin_ + node.children_.Length - 1
                    ||  node.children_[c - node.cMin_] == null)
                {
                    break;
                }
                node = node.children_[c - node.cMin_];
                if (node.bLitEnd_) matchPos = pos + 1;
            }
            if (matchPos >= _pos)
            {
                _pos= matchPos;
                return true;
            }
            else return false;
        }
        #endregion PEG '<Literal>' / '<Literal>'/i / [low1-high1,low2-high2..] / [<CharList>]
    }
    #endregion Parsers
}
