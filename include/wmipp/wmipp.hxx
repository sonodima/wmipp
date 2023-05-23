/**
 * WMI++ takes away the pain from interfacing with
 * the Windows Management Instrumentation in C++.
 *
 * Author:			sonodima
 * Repo:			https://github.com/sonodima/wmipp
 * C++ Version:		20>
 * Supported OS:	Windows
 *
 * ---------------------------------------------------
 *
 * Copyright © 2023 sonodima
 *
 * Permission is hereby granted, free of charge, to
 * any person obtaining a copy of this software and
 * associated documentation files (the “Software”),
 * to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify,
 * merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom
 * the Software is furnished to do so, subject to the
 * following conditions:
 *
 * The above copyright notice and this permission notice
 * shall be included in all copies or substantial portions
 * of the Software.
 *
 * THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF
 * ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED
 * TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
 * PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT
 * SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR
 * ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
 * ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
 * OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 */

#ifndef SD_WMIPP_HXX
#define SD_WMIPP_HXX

#include <memory>
#include <optional>
#include <utility>
#include <stdexcept>
#include <vector>
#include <string>
#include <string_view>
#include <concepts>

#include <atlsafe.h>
#include <comdef.h>
#include <Wbemidl.h>

#pragma comment(lib, "wbemuuid.lib")

namespace wmipp
{
	class Interface;

	struct Exception final : std::runtime_error{
		explicit Exception(const std::string& message) : std::runtime_error(message) {}
	};

	template <class T>
	[[nodiscard]] std::optional<T> ConvertVariant(const CComVariant& variant) {
		std::optional<T> result = std::nullopt;
		try { result = static_cast<T>(variant_t(variant)); }
		catch (...) { }
		return result;
	}

	template <class T>
		requires std::is_same_v<T, std::string> || std::is_same_v<T, std::wstring>
	[[nodiscard]] std::optional<T> ConvertVariant(const CComVariant& variant) {
		// By converting the variant into a bstr_t first, we can automatically
		// handle character type conversions and support std::string and std::wstring.
		std::optional<T> result = std::nullopt;
		if (const auto temp = ConvertVariant<bstr_t>(variant)) {
			result = static_cast<T>(*temp);
		}

		return result;
	}

	template <class T>
		requires std::is_same_v<T, std::vector<typename T::value_type>>
	[[nodiscard]] std::optional<T> ConvertVariant(const CComVariant& variant) {
		// Allocate a temporary safe array object to read the data from the variant.
		// This allows for automatic type conversions and other QOL improvements.
		CComSafeArray<typename T::value_type> safe_array;
		try { safe_array.Attach(variant.parray); }
		catch (...) { return std::nullopt; }

		// Copy the data from the safe array into a vector, element by element.
		T result{};
		result.reserve(safe_array.GetCount());
		for (unsigned long i = 0; i < safe_array.GetCount(); ++i) {
			result.push_back(safe_array.GetAt(i));
		}

		safe_array.Detach();
		return result;
	}

	template <class T>
		requires std::is_same_v<T, std::vector<typename T::value_type>>
		&& (std::is_same_v<typename T::value_type, std::string>
			|| std::is_same_v<typename T::value_type, std::wstring>)
	[[nodiscard]] std::optional<T> ConvertVariant(const CComVariant& variant) {
		// Read all data as a vector of BSTRs first.
		// We will convert each BSTR into the desired string type later.
		const auto intm = ConvertVariant<std::vector<BSTR>>(variant);
		if (!intm) return std::nullopt;

		T result{};
		result.reserve(intm->size());
		for (const auto& element : *intm) {
			// Convert it into a bstr_t first to automatically handle character type conversions.
			const auto temp = static_cast<typename T::value_type>(bstr_t(element));
			result.emplace_back(temp);
			SysFreeString(element);
		}

		return result;
	}

	/**
	 * \brief This class encapsulates a WMI object obtained from a query result.
	 * It provides a convenient interface to access its properties.
	 */
	class Object{
		friend class QueryResult;

	protected:
		Object(std::shared_ptr<const Interface> iface, CComPtr<IWbemClassObject> object)
			: iface_(std::move(iface)), object_(std::move(object)) {}

	public:
		/**
		 * \brief This function retrieves the value of the property with the given name from the WMI object.
		 * The value is converted to the specified type, and if the conversion fails, std::nullopt is returned.
		 * \tparam T The type of the property value to retrieve.
		 * \param name The name of the property to retrieve, specified as a wide string view.
		 * \return std::optional containing the retrieved property value, or std::nullopt if retrieval fails.
		 * \note Certain type conversions may throw asserts in debug mode if the conversion is not possible.
		 */
		template <class T = variant_t>
		[[nodiscard]] std::optional<T> GetProperty(const std::wstring_view name) const {
			CComVariant variant;
			const auto result = object_->Get(
				name.data(),
				0,
				&variant,
				nullptr,
				nullptr);
			if (FAILED(result)) {
				return std::nullopt;
			}

			// Only perform the variant type conversion if a return type other than variant_t
			// is specified.
			if constexpr (std::is_same_v<T, variant_t>) {
				return variant;
			}

			return ConvertVariant<T>(variant);
		}

		/**
		 * \brief Equality operator overload to compare Objects.
		 * Comparison ignores the objects source (the server and the namespace they
		 * were retrieved from) and all qualifiers.
		 * \param other The Object to compare with.
		 * \return true if the Objects are equal, false otherwise.
		 */
		bool operator==(const Object& other) const {
			if (object_ == nullptr && other.object_ == nullptr) return true;
			if (object_ == nullptr || other.object_ == nullptr) return false;
			return object_->CompareTo(
				WBEM_FLAG_IGNORE_OBJECT_SOURCE | WBEM_FLAG_IGNORE_QUALIFIERS,
				other.object_) == WBEM_S_SAME;
		}

		/**
		 * \brief Inequality operator overload to compare Objects.
 		 * Comparison ignores the objects source (the server and the namespace they
		 * were retrieved from) and all qualifiers.
		 * \param other The Object to compare with.
		 * \return true if the Objects are not equal, false otherwise.
		 */
		bool operator!=(const Object& other) const {
			return !(*this == other);
		}

	private:
		std::shared_ptr<const Interface> iface_;
		CComPtr<IWbemClassObject> object_;
	};

	/**
	 * \brief Encapsulates a collection of objects obtained from a query operation. 
	 * It provides methods to access and retrieve properties from the objects in a convenient manner.
	 * The class supports iterating over the objects using range-based for loops.
	 */
	class QueryResult{
		friend class Interface;

	protected:
		QueryResult(std::shared_ptr<const Interface> iface, const CComPtr<IEnumWbemClassObject>& enumerator)
				: iface_(std::move(iface)) {
			if (enumerator) PopulateObjects(enumerator);
		}

	public:
		/**
		 * \brief Finds and retrieves the value of a specific property in the populated objects by name.
		 * If the property is found and its type matches the provided template type T, the value is returned.
		 * \see Object::GetProperty for more information.
		 * \tparam T The type of the property to retrieve.
		 * \param name The name of the property to retrieve.
		 * \return An optional value containing the property value if found, or an empty optional if not found.
		 */
		template <class T = variant_t>
		[[nodiscard]] std::optional<T> GetProperty(const std::wstring_view name) const {
			for (const Object& obj : *this) {
				if (auto value = obj.GetProperty<T>(name)) {
					return value;
				}
			}

			return std::nullopt;
		}

		/**
		 * \brief Retrieves the value of a specific property in the object at the specified index by name.
		 * If the property is found and its type matches the provided template type T, the value is returned.
		 * \see Object::GetProperty for more information.
		 * \tparam T The type of the property to retrieve.
		 * \param name The name of the property to retrieve.
		 * \param index The index of the object in the objects vector to retrieve the property from.
		 * \return An optional value containing the property value exists, or an empty optional if not found.
		 */
		template <class T = variant_t>
		[[nodiscard]] std::optional<T> GetProperty(const std::wstring_view name, const std::size_t index) const {
			if (index >= Count()) return std::nullopt;
			return GetAt(index).GetProperty<T>(name);
		}

		/**
		 * \brief Returns the number of objects in the result.
		 */
		[[nodiscard]] std::size_t Count() const {
			return objects_.size();
		}

		/**
		 * \param index The index of the object to access.
		 * \return A const reference to the object at the specified index.
		 * \throws std::out_of_range if the index is out of range.
		 */
		[[nodiscard]] const Object& GetAt(const std::size_t index) const {
			return objects_.at(index);
		}

		/**
		 * \param index The index of the object to access.
		 * \return A const reference to the object at the specified index.
		 * \throws std::out_of_range if the index is out of range.
		 */
		const Object& operator[](const std::size_t index) const {
			return GetAt(index);
		}

		[[nodiscard]] std::vector<Object>::const_iterator begin() const {
			return objects_.begin();
		}

		[[nodiscard]] std::vector<Object>::const_iterator end() const {
			return objects_.end();
		}

	private:
		std::shared_ptr<const Interface> iface_;
		std::vector<Object> objects_;

		/**
		 * \brief Fills the objects vector by iterating over the given enumerator.
		 * We do this here so that we don't need to clone the enumerator and wrap it into it's own ::enumerator.
		 * Each retrieved IWbemClassObject is wrapped in an Object and added to the objects vector.
		 * \param enumerator A pointer to the IEnumWbemClassObject enumerator.
		 */
		void PopulateObjects(const CComPtr<IEnumWbemClassObject>& enumerator) {
			while (true) {
				ULONG returned_count = 0;
				CComPtr<IWbemClassObject> object = nullptr;
				const auto result = enumerator->Next(
					WBEM_INFINITE,
					1,
					&object,
					&returned_count);
				if (FAILED(result) || returned_count == 0) {
					break;
				}

				objects_.emplace_back(Object(iface_, object));
			}
		}
	};

	/**
	 * \brief Manages a connection to the WMI service and provides a convenient interface
	 * to query WMI objects.
	 * This class also handles COM initialization and cleanup.
	 */
	class Interface : public std::enable_shared_from_this<const Interface> {
	public:
		/**
		 * Initializes the COM library and creates a connection to the WMI service.
		 * \param path The path to the WMI namespace to connect to.
		 * \throws wmipp::Exception if the COM library fails to initialize or the connection
		 * to the WMI service fails.
		 */
		static std::shared_ptr<Interface> Create(std::string_view path = "cimv2");

		Interface(const Interface& other) = default;
		Interface& operator=(const Interface& other) = default;

		Interface(Interface&& other) noexcept = delete;
		Interface& operator=(Interface&& other) noexcept = delete;

		/**
		 * \brief Executes a WQL query and returns the result.
		 * \param query The WQL query to execute.
		 * \return A QueryResult instance containing the result of the query.
		 * \throws wmipp::Exception if the query fails to execute.
		 */
		[[nodiscard]] QueryResult ExecuteQuery(const std::wstring_view query) const {
			CComPtr<IEnumWbemClassObject> enumerator;
			const auto result = services_->ExecQuery(
				bstr_t("WQL"),
				bstr_t(query.data()),
				WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
				nullptr,
				&enumerator);
			if (FAILED(result)) {
				throw Exception("Failed to execute WQL query");
			}

			return {shared_from_this(), enumerator};
		}

	private:
		struct MakeSharedEnabler;

		CComPtr<IWbemLocator> locator_;
		CComPtr<IWbemServices> services_;

		explicit Interface(const std::string_view path) {
			auto result = CoInitializeEx(nullptr, 0);
			if (FAILED(result)) throw Exception("Failed to initialize the COM library");

			result = CoCreateInstance(
				CLSID_WbemLocator,
				nullptr,
				CLSCTX_INPROC_SERVER,
				IID_IWbemLocator,
				reinterpret_cast<LPVOID*>(&locator_));
			if (FAILED(result)) {
				CoUninitialize();
				throw Exception("Failed to create WbemLocator object");
			}

			result = locator_->ConnectServer(
				bstr_t(R"(\\.\root\)") + bstr_t(path.data()),
				nullptr,
				nullptr,
				nullptr,
				0,
				nullptr,
				nullptr,
				&services_);
			if (FAILED(result)) {
				CoUninitialize();
				throw Exception("Could not connect to WMI service");
			}

			result = CoSetProxyBlanket(
				services_,
				RPC_C_AUTHN_DEFAULT,
				RPC_C_AUTHZ_NONE,
				COLE_DEFAULT_PRINCIPAL,
				RPC_C_AUTHN_LEVEL_DEFAULT,
				RPC_C_IMP_LEVEL_IMPERSONATE,
				nullptr,
				EOAC_NONE);
			if (FAILED(result)) {
				CoUninitialize();
				throw Exception("Could not set proxy blanket");
			}
		}

		/**
		 * \brief Uninitializes the COM library and releases the WMI service connection.
		 * \note This function should never be called manually, as the lifetime of the
		 * instances of this class is automatically managed by std::shared_ptr.
		 */
		~Interface() {
			if (services_) services_.Release();
			if (locator_) locator_.Release();
			CoUninitialize();
		}
	};

	/**
	 * \see https://stackoverflow.com/questions/8147027/how-do-i-call-stdmake-shared-on-a-class-with-only-protected-or-private-const/8147213#8147213
	 */
	struct Interface::MakeSharedEnabler : Interface {
		template <typename... Args>
		explicit MakeSharedEnabler(Args &&... args) : Interface(std::forward<Args>(args)...) {}
	};

	inline std::shared_ptr<Interface> Interface::Create(const std::string_view path) {
		return std::make_shared<MakeSharedEnabler>(path);
	}
} // namespace wmipp

#endif // SD_WMIPP_HXX
