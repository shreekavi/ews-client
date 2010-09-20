require 'rubygems'
require 'nokogiri'
require 'ews_connection'
class EwsService
	attr_accessor :errors
	@@ews_user_name = "room1"
	@@ews_user_password = "Shell123"
	@@ews_endpoint = "https://shell.shellnetworks.com/ews/exchange.asmx"
	def initialize
		@errors = []
	end
	
	#Get a Calendar and show all the Ids and Changekeys in the calendar
	def get_calendar_xml
		#Prepare a SOAP request object and send it to the Server
		#Get the Response, parse the resonse and render the Ids and ChangeKeys
		connection = EwsConnection.new(@@ews_user_name, @@ews_user_password, @@ews_endpoint)
		
		begin
			response_doc = Nokogiri::XML(connection.connect(prepare_calendar_request))
			status = response_doc.xpath(response_code_search_path, "m"=>messages).text
			puts status
						
		rescue Exception => e
			#@errors << "Ooops..., there was an XML exception: #{e}."
		    return false
		end
		
	  	if status.to_s != "NoError"
	  		response_code = response_doc.xpath(response_code_search_path, "m"=>messages).text
	      	return false
	    end
	    return response_doc
	end
	
	## Prepare a SOAP Request Format and obtain the XML
	def prepare_calendar_request
	builder = Nokogiri::XML::Builder.new do |xml|

      xml.send(:"soap:Envelope", xmlns_soap, xmlns_t) {
      	xml.send(:"soap:Body") {
        	xml.send(:"FindItem", xmlns_t, xmlns_messages, 'Traversal'=>'Shallow') {
         			 xml.ItemShape {
		  	    	 	xml.send(:'t:BaseShape') {xml.text "IdOnly"}
	    				}         		
				     xml.ParentFolderIds{
	    				xml.send(:"t:DistinguishedFolderId", 'Id'=>"calendar"){
	    				}
	     			}						
          		}
       		}
    	}
    
	    end
	    return builder.to_xml
	
	end
	
	def parse_calendar_for_meetings(xml_doc)
		## Obtain the response and transform that into XML

		## Use Xpath to iterate through calendar Items and fetch Ids and Keys
		xml_doc.xpath("//t:Items/t:CalendarItem/t:ItemId" , "t" => types).each do |node|
		 puts node['Id']
		 puts node['ChangeKey']
			 #Create an array of hashes. Each has containing Id and ChangeKey Key-Value pair
			 #@calendar_response_doc= get_calendar_items(node['Id'], node['ChangeKey'])
			 get_each_calendar_item(@calendar_response_doc)		
		end
	end
	
	## Get Calendar Items
	def get_calendar_items(id, change_key)
		builder = Nokogiri::XML::Builder.new do |xml|

    	  xml.send(:"soap:Envelope",xmlns_soap, xmlns_t) {
    	    xml.send(:"soap:Body") {
    	     xml.send(:"GetItem", messages) {
			    xml.ItemShape {
			     xml.send(:'t:BaseShape') {xml.text "IdOnly"}       		
			     xml.send(:'t:AdditionalProperties'){
			    	xml.send(:'t:FieldURI', "FieldURI"=>"item:Subject") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:Start") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:End") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:Location") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:When") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:Resources") {}
			    	xml.send(:'t:FieldURI', "FieldURI"=>"calendar:MeetingTimeZone") {}
		        	}
	     		}
	   			xml.ItemIds {
	        		xml.send(:'t:ItemId', "Id"=>id, "ChangeKey"=>change_key) {}
	   			}  						
          	}
       	}
    }
		end
  	return builder.to_xml
	
	end
	
	def get_each_calendar_item(cal_response_doc)
		xml_doc = Nokogiri::XML(cal_response_doc)	
		xml_doc.xpath("//m:Items/t:CalendarItem/*" , "t" =>types,"m"=>messages).each do |d|
			puts d.text
		end
	
	end
	
  def xmlns_xsi
    {'xmlns:xsi' => "http://www.w3.org/2001/XMLSchema-instance"}
  end
  
  def xmlns_xsd
    {'xmlns:xsd' => "http://www.w3.org/2001/XMLSchema"}
  end
  
  def xmlns_soap
    {'xmlns:soap' => "http://schemas.xmlsoap.org/soap/envelope/"}
  end
  
  def xmlns_t
    {'xmlns:t' => "http://schemas.microsoft.com/exchange/services/2006/types"}
  end
  
  def xmlns_messages
    {'xmlns' => "http://schemas.microsoft.com/exchange/services/2006/messages"}
  end
  
  def xmlns_types
    {'xmlns' => "http://schemas.microsoft.com/exchange/services/2006/types"}
  end
  
  def types
	  "http://schemas.microsoft.com/exchange/services/2006/types"
  end
  
  def messages
  	"http://schemas.microsoft.com/exchange/services/2006/messages"
  end
  
  def response_code_search_path
  	"//m:ResponseMessages/m:FindItemResponseMessage/m:ResponseCode"
  end

end
