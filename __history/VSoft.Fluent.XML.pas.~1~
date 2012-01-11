{***************************************************************************}
{                                                                           }
{           VSoft.Fluent.XML                                                }
{                                                                           }
{           Copyright (C) 2011 Vincent Parrett                              }
{                                                                           }
{           http://www.finalbuilder.com                                     }
{                                                                           }
{                                                                           }
{***************************************************************************}
{                                                                           }
{  Licensed under the Apache License, Version 2.0 (the "License");          }
{  you may not use this file except in compliance with the License.         }
{  You may obtain a copy of the License at                                  }
{                                                                           }
{      http://www.apache.org/licenses/LICENSE-2.0                           }
{                                                                           }
{  Unless required by applicable law or agreed to in writing, software      }
{  distributed under the License is distributed on an "AS IS" BASIS,        }
{  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. }
{  See the License for the specific language governing permissions and      }
{  limitations under the License.                                           }
{                                                                           }
{***************************************************************************}

/// this is a based on the work of  GpFluentXML 
//  http://17slon.com/blogs/gabr/files/GpFluentXml.pas
//
// it uses msxml 6 rather than omnixml
//
// NOT FULLY TESTED YET!!!


unit VSoft.Fluent.XML;

interface

uses
  VSoft.MSXML.6;

type
  IFluentXmlBuilder = interface
  ['{91F596A3-F5E3-451C-A6B9-C5FF3F23ECCC}']
    function  GetXml: IXMLDOMDocument;
    function  AddChild(const name: string): IFluentXmlBuilder; overload;
    function  AddChild(const name: string; value: Variant): IFluentXmlBuilder; overload;
    function  AddComment(const comment: string): IFluentXmlBuilder;
    function  AddProcessingInstruction(const target, data: string): IFluentXmlBuilder;
    function  AddSibling(const name: string): IFluentXmlBuilder; overload;
    function  AddSibling(const name: string; value: Variant): IFluentXmlBuilder; overload;
    function  Anchor(var node: IXMLDOMNode): IFluentXmlBuilder;
    function  Mark: IFluentXmlBuilder;
    function  Return: IFluentXmlBuilder;
    function  SetAttrib(const name, value: string): IFluentXmlBuilder;
    function  Up: IFluentXmlBuilder;
    property Attrib[const name, value: string]: IFluentXmlBuilder read SetAttrib; default;
    property Xml: IXMLDOMDocument read GetXml;
  end;

function CreateFluentXml: IFluentXmlBuilder;

implementation

uses
  SysUtils,
  Variants,
  Classes;

const
  DEFAULT_DECIMALSEPARATOR  = '.'; // don't change!
  DEFAULT_TRUE              = '1'; // don't change!
  DEFAULT_TRUE_STR          = 'true'; // don't change!
  DEFAULT_FALSE             = '0'; // don't change!
  DEFAULT_FALSE_STR         = 'false'; // don't change!
  DEFAULT_DATETIMESEPARATOR = 'T'; // don't change!
  DEFAULT_DATESEPARATOR     = '-'; // don't change!
  DEFAULT_TIMESEPARATOR     = ':'; // don't change!
  DEFAULT_MSSEPARATOR       = '.'; // don't change!


type
  TFluentXmlBuilder = class(TInterfacedObject, IFluentXmlBuilder)
  private
    FxbActiveNode : IXMLDOMNode;
    FxbMarkedNodes: IInterfaceList;
    FxbXmlDoc     : IXMLDOMDocument;
  protected
    function ActiveNode: IXMLDOMNode;
  protected
    function GetXml: IXMLDOMDocument;
  public
    constructor Create;
    destructor  Destroy; override;
    function  AddChild(const name: string): IFluentXmlBuilder; overload;
    function  AddChild(const name: string; value: Variant): IFluentXmlBuilder; overload;
    function  AddComment(const comment: string): IFluentXmlBuilder;
    function  AddProcessingInstruction(const target, data: string): IFluentXmlBuilder;
    function  AddSibling(const name: string): IFluentXmlBuilder; overload;
    function  AddSibling(const name: string; value: Variant): IFluentXmlBuilder; overload;
    function  Anchor(var node: IXMLDOMNode): IFluentXmlBuilder;
    function  Mark: IFluentXmlBuilder;
    function  Return: IFluentXmlBuilder;
    function  SetAttrib(const name, value: string): IFluentXmlBuilder;
    function  Up: IFluentXmlBuilder;
  end; { TFluentXmlBuilder }

{ globals }

function CreateFluentXml: IFluentXmlBuilder;
begin
  Result := TFluentXmlBuilder.Create;
end; { CreateFluentXml }

{ TFluentXmlBuilder }

constructor TFluentXmlBuilder.Create;
begin
  inherited Create;
  FxbXmlDoc := CoDOMDocument60.Create;
  fxbMarkedNodes := TInterfaceList.Create;
end; { TFluentXmlBuilder.Create }

destructor TFluentXmlBuilder.Destroy;
begin
  if fxbMarkedNodes.Count > 0 then
    raise Exception.Create('''Mark'' stack is not completely empty');
  inherited;
end; { TFluentXmlBuilder.Destroy }

function TFluentXmlBuilder.ActiveNode: IXMLDOMNode;
begin
  if assigned(fxbActiveNode) then
    Result := fxbActiveNode
  else begin
    Result := FxbXmlDoc.documentElement;
    if not assigned(Result) then
      Result := fxbXmlDoc;
  end;
end; { TFluentXmlBuilder.ActiveNode }

function TFluentXmlBuilder.AddChild(const name: string): IFluentXmlBuilder;
var
  parentNode : IXMLDOMNode;
begin
  parentNode := ActiveNode;
  FxbActiveNode := FxbXmlDoc.createElement(name);
  parentNode.appendChild(FxbActiveNode);
  Result := Self;
end; { TFluentXmlBuilder.AddChild }

function SetTextChild(node: IXMLDOMNode; value: string): IXMLDOMNode;
var
  iText: integer;
begin
  iText := 0;
  while iText < node.ChildNodes.Length do
  begin
    if node.ChildNodes.Item[iText].NodeType = NODE_TEXT then
      node.RemoveChild(node.ChildNodes.Item[iText])
    else
      Inc(iText);
  end; //while
  Result := node.ownerDocument.CreateTextNode(value);
  node.AppendChild(Result);
end; { SetTextChild }

function XMLExtendedToStr(value: extended): string;
begin
  Result := StringReplace(FloatToStr(value),
    DecimalSeparator, DEFAULT_DECIMALSEPARATOR, [rfReplaceAll]);
end; { XMLExtendedToStr }

function XMLCurrencyToStr(value: Currency): string;
begin
  Result := StringReplace(CurrToStr(value),
    DecimalSeparator, DEFAULT_DECIMALSEPARATOR, [rfReplaceAll]);
end; { XMLExtendedToStr }


function XMLBoolToStr(value: boolean; useBoolStrs: boolean = false): string;
begin
  if value then
    if useBoolStrs then
      Result := DEFAULT_TRUE_STR
    else
      Result := DEFAULT_TRUE
  else
    if useBoolStrs then
      Result := DEFAULT_FALSE_STR
    else
      Result := DEFAULT_FALSE;
end; { XMLBoolToStr }

function XMLDateTimeToStr(value: TDateTime): string;
begin
  if Trunc(value) = 0 then
    Result := FormatDateTime('"'+DEFAULT_DATETIMESEPARATOR+'"hh":"mm":"ss.zzz',value)
  else
    Result := FormatDateTime('yyyy-mm-dd"'+
      DEFAULT_DATETIMESEPARATOR+'"hh":"mm":"ss.zzz',value);
end; { XMLDateTimeToStr }

function XMLTimeToStr(value: TDateTime): string;
begin
  Result := FormatDateTime('hh":"mm":"ss.zzz',value);
end; { XMLTimeToStr }

function XMLDateToStr(value: TDateTime): string;
begin
  Result := FormatDateTime('yyyy-mm-dd',value);
end; { XMLDateToStr }


function XMLDateTimeToStrEx(value: TDateTime): string;
begin
  if Trunc(value) = 0 then
    Result := XMLTimeToStr(value)
  else if Frac(Value) = 0 then
    Result := XMLDateToStr(value)
  else
    Result := XMLDateTimeToStr(value);
end; { XMLDateTimeToStrEx }




function XMLVariantToStr(value: Variant): string;
begin
  case VarType(value) of
    varSingle, varDouble, varCurrency:
      Result := XMLExtendedToStr(value);
    varDate:
      Result := XMLDateTimeToStrEx(value);
    varBoolean:
      Result := XMLBoolToStr(value);
    else
      Result := value;
  end; //case
end; { XMLVariantToStr }


function TFluentXmlBuilder.AddChild(const name: string; value: Variant): IFluentXmlBuilder;
begin
  Result := AddChild(name);
  SetTextChild(fxbActiveNode, XMLVariantToStr(value));
end; { TFluentXmlBuilder.AddChild }

function TFluentXmlBuilder.AddComment(const comment: string): IFluentXmlBuilder;
begin
  ActiveNode.AppendChild(fxbXmlDoc.CreateComment(comment));
  Result := Self;
end; { TFluentXmlBuilder.AddComment }

function TFluentXmlBuilder.AddProcessingInstruction(const target, data: string):
  IFluentXmlBuilder;
begin
  ActiveNode.AppendChild(fxbXmlDoc.CreateProcessingInstruction(target, data));
  Result := Self;
end; { TFluentXmlBuilder.AddProcessingInstruction }

function TFluentXmlBuilder.AddSibling(const name: string;
  value: Variant): IFluentXmlBuilder;
begin
  Result := AddSibling(name);
  SetTextChild(fxbActiveNode, XMLVariantToStr(value));
end; { TFluentXmlBuilder.AddSibling }

function TFluentXmlBuilder.AddSibling(const name: string): IFluentXmlBuilder;
var
  parentNode : IXMLDOMNode;
begin
  Result := Up;
  parentNode := ActiveNode;

  FxbActiveNode := FxbXmlDoc.createElement(name);
  parentNode.appendChild(FxbActiveNode);
end; { TFluentXmlBuilder.AddSibling }

function TFluentXmlBuilder.Anchor(var node: IXMLDOMNode): IFluentXmlBuilder;
begin
  node := fxbActiveNode;
  Result := Self;
end; { TFluentXmlBuilder.Anchor }

function TFluentXmlBuilder.GetXml: IXMLDOMDocument;
begin
  Result := fxbXmlDoc;
end; { TFluentXmlBuilder.GetXml }

function TFluentXmlBuilder.Mark: IFluentXmlBuilder;
begin
  fxbMarkedNodes.Add(ActiveNode);
  Result := Self;
end; { TFluentXmlBuilder.Mark }

function TFluentXmlBuilder.Return: IFluentXmlBuilder;
begin
  fxbActiveNode := fxbMarkedNodes.Last as IXMLDOMNode;
  fxbMarkedNodes.Delete(fxbMarkedNodes.Count - 1);
  Result := Self;
end; { TFluentXmlBuilder.Return }

function TFluentXmlBuilder.SetAttrib(const name, value: string): IFluentXmlBuilder;
begin
  if ActiveNode.nodeType = NODE_ELEMENT then
    IXMLDOMElement(ActiveNode).setAttribute(name,value);
  Result := Self;
end; { TFluentXmlBuilder.SetAttrib }

function TFluentXmlBuilder.Up: IFluentXmlBuilder;
begin
  if not assigned(fxbActiveNode) then
    raise Exception.Create('Cannot access a parent at the root level')
  else if fxbActiveNode = fxbXmlDoc.documentElement then
    raise Exception.Create('Cannot create a parent at the document element level')
  else
    fxbActiveNode := ActiveNode.ParentNode;
  Result := Self;
end; { TFluentXmlBuilder.Up }

end.

